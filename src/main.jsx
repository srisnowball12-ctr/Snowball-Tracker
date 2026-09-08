import React, { useEffect, useMemo, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'
import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'
import {
  Upload, Download, FileText, LogOut, RefreshCw,
  LayoutDashboard, Table2, Settings, Users,
  Eye, EyeOff, ArrowLeft, UserPlus, UserMinus,
  ShieldCheck, CheckCircle2, Bell, X, Mail
} from 'lucide-react'
import logoUrl from './logo.png'
import './styles.css'

/*
  FINAL CONSOLIDATED BUILD
  - Excel calendar dates are preserved without timezone shifting.
  - Consolidated uploads replace the existing dataset dates before inserting
    every Excel row.
  - SWP detection ignores Folio and uses Investor + Scheme.
  - Employee Red/SWP labels are treated as inputs; the recurring SWP pattern
    is the final authority for Red vs SWP.
  - Switch and STP remain fixed classifications.
  - Duplicate header Logout is removed; sidebar Logout remains.
*/

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
)

const money = n =>
  new Intl.NumberFormat('en-IN', {
    style: 'currency',
    currency: 'INR',
    maximumFractionDigits: 0
  }).format(Number(n || 0))

const compactMoney = n => {
  const value = Number(n || 0)
  if (value >= 10000000) return `₹${(value / 10000000).toFixed(2)} Cr`
  if (value >= 100000) return `₹${(value / 100000).toFixed(2)} L`
  return money(value)
}

const norm = s =>
  String(s ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')

/*
  IMPORTANT DATE RULE
  -------------------
  Never convert an Excel calendar date through local time + toISOString().
  Excel dates are read as serial numbers where possible and converted using
  UTC calendar arithmetic. Text dates in DD/MM/YYYY or DD-MM-YYYY are also
  parsed explicitly.
*/
const iso = value => {
  if (value === null || value === undefined || value === '') return null

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return [
      value.getUTCFullYear(),
      String(value.getUTCMonth() + 1).padStart(2, '0'),
      String(value.getUTCDate()).padStart(2, '0')
    ].join('-')
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const excelEpoch = Date.UTC(1899, 11, 30)
    const wholeDays = Math.floor(value)
    const d = new Date(excelEpoch + wholeDays * 86400000)

    return [
      d.getUTCFullYear(),
      String(d.getUTCMonth() + 1).padStart(2, '0'),
      String(d.getUTCDate()).padStart(2, '0')
    ].join('-')
  }

  const text = String(value).trim()

  const isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/)
  if (isoMatch) {
    const yyyy = Number(isoMatch[1])
    const mm = Number(isoMatch[2])
    const dd = Number(isoMatch[3])
    const test = new Date(Date.UTC(yyyy, mm - 1, dd))

    if (
      test.getUTCFullYear() === yyyy &&
      test.getUTCMonth() === mm - 1 &&
      test.getUTCDate() === dd
    ) {
      return `${yyyy}-${String(mm).padStart(2, '0')}-${String(dd).padStart(2, '0')}`
    }
    return null
  }

  const dmy = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/)
  if (dmy) {
    const dd = Number(dmy[1])
    const mm = Number(dmy[2])
    const yyyy = Number(dmy[3])
    const test = new Date(Date.UTC(yyyy, mm - 1, dd))

    if (
      test.getUTCFullYear() === yyyy &&
      test.getUTCMonth() === mm - 1 &&
      test.getUTCDate() === dd
    ) {
      return `${yyyy}-${String(mm).padStart(2, '0')}-${String(dd).padStart(2, '0')}`
    }
    return null
  }

  /*
    Final fallback for strings such as "Aug 25 2026".
    Use local calendar components, not toISOString(), so a timezone cannot
    move the transaction to the previous day.
  */
  const fallback = new Date(text)
  if (Number.isNaN(fallback.getTime())) return null

  return [
    fallback.getFullYear(),
    String(fallback.getMonth() + 1).padStart(2, '0'),
    String(fallback.getDate()).padStart(2, '0')
  ].join('-')
}

function sourceLabel(value) {
  const text = String(value || '').trim().toLowerCase()

  if (
    text.includes('swp') ||
    text.includes('systematic withdrawal')
  ) return 'SWP'

  if (
    text.includes('switch') ||
    text.includes('switch transaction') ||
    text.includes('switch-in') ||
    text.includes('switch-out')
  ) return 'Switch'

  if (
    text.includes('stp') ||
    text.includes('systematic transfer')
  ) return 'STP'

  if (
    text.includes('red') ||
    text.includes('redeem') ||
    text.includes('redemption')
  ) return 'Redemption'

  return value || 'Redemption'
}

function explicitClassification(value) {
  const label = sourceLabel(value)

  if (['Switch', 'STP'].includes(label)) return label

  /*
    SWP is intentionally NOT treated as final here.
    Employee-entered SWP and employee-entered Redemption are both tested
    against the recurring SWP pattern.
  */
  return null
}

/*
  FINAL SWP CLASSIFICATION LOGIC
  ------------------------------
  1. Folio number is completely ignored.
  2. Group by Investor + Scheme only.
  3. Switch and STP are always retained as their own classifications.
  4. Both employee Red/Redemption and employee SWP rows are candidates.
  5. A qualifying SWP sequence contains at least 3 eligible transactions.
  6. Transactions must be consecutive eligible withdrawals in date order.
  7. Each consecutive gap must be 25–40 days.
  8. Successive amounts must be within 15%.
  9. If a qualifying sequence is found, every transaction in that sequence
     is classified as SWP, irrespective of the employee source label.
  10. Otherwise the transaction remains Redemption.

  "Consecutive" is strict: a gap below 25 days or above 40 days breaks the
  sequence. We do not skip an intervening candidate transaction.
*/
function classifyRows(rows) {
  const output = rows.map(row => ({ ...row }))
  const groups = new Map()

  output.forEach(row => {
    const key = [
      norm(row.investor_name),
      norm(row.scheme)
    ].join('|')

    if (!groups.has(key)) groups.set(key, [])
    groups.get(key).push(row)
  })

  groups.forEach(items => {
    items.sort((a, b) =>
      String(a.transaction_date || '').localeCompare(
        String(b.transaction_date || '')
      )
    )

    items.forEach(row => {
      const label = sourceLabel(row.original_transaction_type)

      if (label === 'Switch' || label === 'STP') {
        row.display_classification = label
      } else {
        row.display_classification = 'Redemption'
      }
    })

    const candidates = items
      .filter(row => {
        const label = sourceLabel(row.original_transaction_type)
        return (
          (label === 'Redemption' || label === 'SWP') &&
          row.transaction_date &&
          Number.isFinite(Number(row.amount)) &&
          Number(row.amount) > 0
        )
      })
      .map(row => ({
        row,
        date: new Date(`${row.transaction_date}T00:00:00`),
        amount: Number(row.amount)
      }))
      .filter(item => !Number.isNaN(item.date.getTime()))

    for (let start = 0; start <= candidates.length - 3; start++) {
      const sequence = [candidates[start]]

      for (let next = start + 1; next < candidates.length; next++) {
        const previous = sequence[sequence.length - 1]
        const current = candidates[next]

        const days = Math.round(
          (current.date - previous.date) / 86400000
        )

        /*
          Strict consecutive pattern:
          any candidate inside the sequence must itself be 25–40 days
          from the previous candidate.
        */
        if (days < 25 || days > 40) break

        const amountDiff =
          Math.abs(current.amount - previous.amount) /
          Math.max(previous.amount, current.amount, 1)

        if (amountDiff > 0.15) break

        sequence.push(current)
      }

      if (sequence.length >= 3) {
        sequence.forEach(item => {
          item.row.display_classification = 'SWP'
        })
      }
    }
  })

  return output
}

function mapRow(row) {
  const lookup = Object.fromEntries(
    Object.entries(row).map(([k, v]) => [norm(k), v])
  )

  const get = (...keys) =>
    keys
      .map(k => lookup[norm(k)])
      .find(v =>
        v !== undefined &&
        v !== null &&
        v !== ''
      )

  const amountRaw = get(
    'Amount(₹)',
    'Amount',
    'amount',
    'Transaction Amount'
  )

  const amount =
    typeof amountRaw === 'number'
      ? amountRaw
      : Number(
          String(amountRaw || '')
            .replace(/[₹,\s]/g, '')
        )

  const originalType =
    get(
      'Type',
      'Transaction Type',
      'Transaction Type Description',
      'Txn Type',
      'Txn Type Description',
      'Nature of Transaction',
      'Transaction Nature',
      'Transaction Description',
      'Description',
      'Remarks',
      'Source',
      'original_transaction_type'
    ) || null

  return {
    rm_name: get(
      'Partner/Employee',
      'Partner',
      'Employee',
      'RM',
      'rm_name'
    ) || null,

    group_name:
      get('Group', 'group_name') || null,

    investor_name: get(
      'Investor',
      'Investor Name',
      'Client Name',
      'investor_name'
    ) || null,

    transaction_date: iso(
      get(
        'Date',
        'Redemption Date',
        'Transaction Date',
        'transaction_date'
      )
    ),

    folio_no: String(
      get(
        'Folio No/Demat A/C',
        'Folio No',
        'Folio',
        'folio_no'
      ) || ''
    ) || null,

    scheme:
      get('Scheme', 'Fund', 'scheme') || null,

    amount:
      Number.isFinite(amount)
        ? amount
        : null,

    original_transaction_type:
      sourceLabel(originalType),

    classified_transaction_type:
      sourceLabel(originalType) === 'Switch'
        ? 'Switch'
        : sourceLabel(originalType) === 'STP'
          ? 'STP'
          : 'Redemption',

    classification_status: 'Completed',
    classification_reason: null
  }
}

function App() {
  const [session, setSession] = useState(null)
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [showPassword, setShowPassword] = useState(false)

  const [forgotMode, setForgotMode] = useState(false)
  const [forgotEmail, setForgotEmail] = useState('')
  const [forgotMessage, setForgotMessage] = useState('')

  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)

  const [page, setPage] = useState('dashboard')
  const [rows, setRows] = useState([])
  const [rms, setRms] = useState([])

  const [rm, setRm] = useState('All')
  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')
  const [period, setPeriod] = useState('YTD')

  const [uploading, setUploading] = useState(false)
  const [message, setMessage] = useState('')

  const [newRm, setNewRm] = useState('')
  const [settingsTab, setSettingsTab] = useState('current')

  const [adminModal, setAdminModal] = useState(null)
  const [adminEmail, setAdminEmail] = useState('')

  const [notifications, setNotifications] = useState([])
  const [showNotifications, setShowNotifications] = useState(false)

  const [currentPage, setCurrentPage] = useState(1)
  const pageSize = 50

  useEffect(() => {
    supabase.auth
      .getSession()
      .then(({ data }) => setSession(data.session))

    const {
      data: { subscription }
    } = supabase.auth.onAuthStateChange(
      (_event, newSession) => {
        setSession(newSession)

        if (!newSession) {
          setPage('dashboard')
          setRows([])
          setRms([])
        }
      }
    )

    return () => subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (session) {
      loadData()
      loadRms()
      loadNotifications()
    }
  }, [session])

  useEffect(() => {
    setCurrentPage(1)
  }, [rm, from, to, period])

  async function login(e) {
    e.preventDefault()
    setLoading(true)
    setError('')

    const { error: authError } =
      await supabase.auth.signInWithPassword({
        email,
        password
      })

    if (authError) setError(authError.message)
    setLoading(false)
  }

  async function sendResetPassword(e) {
    e.preventDefault()
    setError('')
    setForgotMessage('')
    setLoading(true)

    const { error: resetError } =
      await supabase.auth.resetPasswordForEmail(
        forgotEmail,
        { redirectTo: window.location.origin }
      )

    if (resetError) {
      setError(resetError.message)
    } else {
      setForgotMessage(
        'Password reset instructions have been sent to your email.'
      )
    }

    setLoading(false)
  }

  async function handleLogout() {
    setError('')
    await supabase.auth.signOut()
    setSession(null)
    setPage('dashboard')
    setRows([])
    setRms([])
    window.location.replace(window.location.origin)
  }

  async function loadData() {
    setLoading(true)
    setError('')

    try {
      const allRows = []
      const batchSize = 1000
      let fromRow = 0

      while (true) {
        const { data, error: readError } =
          await supabase
            .from('transactions')
            .select('*')
            .order('transaction_date', {
              ascending: false
            })
            .order('id', {
              ascending: false
            })
            .range(
              fromRow,
              fromRow + batchSize - 1
            )

        if (readError) throw readError

        const batch = data || []
        allRows.push(...batch)

        if (batch.length < batchSize) break
        fromRow += batchSize
      }

      setRows(allRows)
    } catch (err) {
      setError(
        err?.message ||
        'Unable to load transaction data.'
      )
      setRows([])
    } finally {
      setLoading(false)
    }
  }

  async function loadRms() {
    const { data, error: readError } =
      await supabase
        .from('rms')
        .select('*')
        .order('name')

    if (readError) {
      console.error(readError)
      return
    }

    setRms(data || [])
  }

  async function loadNotifications() {
    const { data, error: readError } =
      await supabase
        .from('notifications')
        .select('*')
        .order('created_at', {
          ascending: false
        })
        .limit(50)

    if (!readError) {
      setNotifications(data || [])
    }
  }

  async function notify(title, body) {
    const item = {
      id: `local-${Date.now()}-${Math.random()}`,
      title,
      body,
      created_at: new Date().toISOString(),
      is_read: false
    }

    setNotifications(prev => [
      item,
      ...prev
    ])

    setMessage(title)

    window.setTimeout(
      () => setMessage(''),
      3500
    )

    const {
      data,
      error: insertError
    } = await supabase
      .from('notifications')
      .insert({
        title,
        body,
        is_read: false
      })
      .select()
      .single()

    if (!insertError && data) {
      setNotifications(prev => [
        data,
        ...prev.filter(
          x => x.id !== item.id
        )
      ])
    }
  }

  async function markNotificationsRead() {
    setNotifications(prev =>
      prev.map(x => ({
        ...x,
        is_read: true
      }))
    )

    await supabase
      .from('notifications')
      .update({ is_read: true })
      .eq('is_read', false)
  }

  async function addRm(e) {
    e.preventDefault()

    const name = newRm.trim()
    if (!name) return

    setError('')

    const { error: insertError } =
      await supabase
        .from('rms')
        .insert({
          name,
          is_active: true,
          is_admin: false
        })

    if (insertError) {
      setError(
        insertError.code === '23505'
          ? 'This RM already exists.'
          : insertError.message
      )
      return
    }

    setNewRm('')

    await notify(
      'RM added',
      `${name} has been added to the organisation.`
    )

    await loadRms()
  }

  async function toggleRmStatus(item) {
    setError('')

    const nextActive =
      item.is_active === false

    if (!item.id) {
      const {
        error: insertError
      } = await supabase
        .from('rms')
        .insert({
          name: item.name,
          is_active: nextActive,
          is_admin: false
        })

      if (
        insertError &&
        insertError.code !== '23505'
      ) {
        setError(insertError.message)
        return
      }
    } else {
      const {
        error: updateError
      } = await supabase
        .from('rms')
        .update({
          is_active: nextActive
        })
        .eq('id', item.id)

      if (updateError) {
        setError(updateError.message)
        return
      }
    }

    await notify(
      nextActive
        ? 'RM activated'
        : 'RM marked inactive',
      `${item.name} has been ${
        nextActive
          ? 'activated'
          : 'marked inactive'
      }.`
    )

    await loadRms()
  }

  function openAdminModal(item) {
    if (item.is_admin) {
      toggleAdmin(
        item,
        item.admin_email || ''
      )
      return
    }

    setAdminModal(item)
    setAdminEmail(
      item.admin_email || ''
    )
  }

  async function toggleAdmin(
    item,
    emailAddress
  ) {
    setError('')

    const nextAdmin = !item.is_admin

    if (!item.id) {
      const {
        error: insertError
      } = await supabase
        .from('rms')
        .insert({
          name: item.name,
          is_active: true,
          is_admin: nextAdmin,
          admin_email:
            nextAdmin
              ? emailAddress
              : null
        })

      if (
        insertError &&
        insertError.code !== '23505'
      ) {
        setError(insertError.message)
        return
      }
    } else {
      const {
        error: updateError
      } = await supabase
        .from('rms')
        .update({
          is_admin: nextAdmin,
          admin_email:
            nextAdmin
              ? emailAddress
              : null
        })
        .eq('id', item.id)

      if (updateError) {
        setError(updateError.message)
        return
      }
    }

    await notify(
      nextAdmin
        ? 'Admin access granted'
        : 'Admin access removed',
      nextAdmin
        ? `${item.name} has been granted admin access for ${emailAddress}.`
        : `${item.name}'s admin access has been removed.`
    )

    setAdminModal(null)
    setAdminEmail('')

    await loadRms()
  }

  async function confirmAdmin(e) {
    e.preventDefault()

    const cleanEmail =
      adminEmail.trim().toLowerCase()

    if (
      !cleanEmail ||
      !/^\S+@\S+\.\S+$/.test(
        cleanEmail
      )
    ) {
      setError(
        'Please enter a valid email address for admin login access.'
      )
      return
    }

    await toggleAdmin(
      adminModal,
      cleanEmail
    )
  }

  /*
    Existing transactions are already classified in Supabase.
    Do not re-run SWP classification on page load.
  */
  const analysedRows = useMemo(
    () =>
      rows.map(row => ({
        ...row,
        display_classification:
          row.classified_transaction_type ||
          'Redemption'
      })),
    [rows]
  )

  const rmDirectory = useMemo(() => {
    const map = new Map()

    analysedRows
      .map(x =>
        String(
          x.rm_name || ''
        ).trim()
      )
      .filter(Boolean)
      .forEach(name => {
        map.set(norm(name), {
          id: null,
          name,
          is_active: true,
          is_admin: false,
          imported: true
        })
      })

    rms.forEach(item => {
      const key = norm(item.name)
      const previous = map.get(key)

      map.set(key, {
        ...(previous || {}),
        ...item,
        name: item.name
      })
    })

    return Array.from(
      map.values()
    ).sort((a, b) =>
      a.name.localeCompare(
        b.name
      )
    )
  }, [analysedRows, rms])

  const activeRmNames =
    useMemo(
      () => [
        'All',
        ...rmDirectory
          .filter(
            x =>
              x.is_active !== false
          )
          .map(x => x.name)
      ],
      [rmDirectory]
    )

  const filtered = useMemo(() => {
    return analysedRows.filter(x => {
      if (
        rm !== 'All' &&
        x.rm_name !== rm
      ) return false

      const d = x.transaction_date

      if (!d) return false
      if (from && d < from) return false
      if (to && d > to) return false

      if (!from && !to) {
        const now = new Date()
        const dt =
          new Date(
            `${d}T00:00:00`
          )

        if (period === 'WTD') {
          const day =
            (now.getDay() + 6) % 7

          const start =
            new Date(now)

          start.setDate(
            now.getDate() - day
          )

          start.setHours(
            0, 0, 0, 0
          )

          if (dt < start) {
            return false
          }
        }

        if (
          period === 'MTD' &&
          (
            dt.getMonth() !==
              now.getMonth() ||
            dt.getFullYear() !==
              now.getFullYear()
          )
        ) {
          return false
        }

        if (period === 'QTD') {
          const q =
            Math.floor(
              now.getMonth() / 3
            )

          if (
            dt.getFullYear() !==
              now.getFullYear() ||
            Math.floor(
              dt.getMonth() / 3
            ) !== q
          ) {
            return false
          }
        }

        if (
          period === 'YTD' &&
          dt.getFullYear() !==
            now.getFullYear()
        ) {
          return false
        }
      }

      return true
    })
  }, [
    analysedRows,
    rm,
    from,
    to,
    period
  ])

  const totals = useMemo(() => {
    const sum = type =>
      filtered
        .filter(
          x =>
            x.display_classification ===
            type
        )
        .reduce(
          (total, x) =>
            total +
            Number(x.amount || 0),
          0
        )

    return {
      Redemption: sum(
        'Redemption'
      ),
      SWP: sum('SWP'),
      Switch: sum('Switch'),
      STP: sum('STP'),
      Investors: new Set(
        filtered
          .map(x =>
            x.investor_name
          )
          .filter(Boolean)
      ).size,
      Transactions:
        filtered.length
    }
  }, [filtered])

  const chartData = useMemo(() => {
    const map = {}

    filtered.forEach(x => {
      const key =
        x.rm_name ||
        'Unassigned'

      map[key] =
        (map[key] || 0) +
        Number(x.amount || 0)
    })

    return Object.entries(map)
      .map(
        ([name, amount]) => ({
          name,
          amount
        })
      )
      .sort(
        (a, b) =>
          b.amount - a.amount
      )
      .slice(0, 10)
  }, [filtered])

  const maxChartValue =
    useMemo(
      () =>
        Math.max(
          ...chartData.map(
            x => x.amount
          ),
          1
        ),
      [chartData]
    )

  const classificationChart =
    useMemo(
      () => [
        {
          name: 'Redemption',
          amount:
            totals.Redemption,
          color: '#c9565a'
        },
        {
          name: 'SWP',
          amount: totals.SWP,
          color: '#5a9b88'
        },
        {
          name: 'Switch',
          amount: totals.Switch,
          color: '#597aa1'
        },
        {
          name: 'STP',
          amount: totals.STP,
          color: '#8067a8'
        }
      ],
      [totals]
    )

  const classificationTotal =
    useMemo(
      () =>
        classificationChart.reduce(
          (sum, x) =>
            sum + x.amount,
          0
        ),
      [classificationChart]
    )

  const donutStyle = useMemo(() => {
    if (!classificationTotal) {
      return {
        background:
          'conic-gradient(#dfe5ed 0 100%)'
      }
    }

    let cursor = 0

    const parts =
      classificationChart
        .filter(x =>
          x.amount > 0
        )
        .map(x => {
          const next =
            cursor +
            (
              x.amount /
              classificationTotal
            ) * 100

          const part =
            `${x.color} ${cursor}% ${next}%`

          cursor = next
          return part
        })

    return {
      background:
        `conic-gradient(${parts.join(', ')})`
    }
  }, [
    classificationChart,
    classificationTotal
  ])

  const monthlyTrend =
    useMemo(() => {
      const map = new Map()

      filtered.forEach(x => {
        const key =
          String(
            x.transaction_date || ''
          ).slice(0, 7)

        if (!key) return

        if (!map.has(key)) {
          map.set(key, {
            Redemption: 0,
            SWP: 0,
            Switch: 0,
            STP: 0
          })
        }

        const bucket =
          map.get(key)

        bucket[
          x.display_classification
        ] =
          (
            bucket[
              x.display_classification
            ] || 0
          ) +
          Number(
            x.amount || 0
          )
      })

      return Array.from(
        map.entries()
      )
        .sort(
          ([a], [b]) =>
            a.localeCompare(b)
        )
        .slice(-6)
        .map(
          ([month, values]) => ({
            month,
            ...values
          })
        )
    }, [filtered])

  const paginatedRows =
    useMemo(() => {
      const start =
        (currentPage - 1) *
        pageSize

      return filtered.slice(
        start,
        start + pageSize
      )
    }, [
      filtered,
      currentPage
    ])

  const totalPages =
    Math.max(
      1,
      Math.ceil(
        filtered.length /
        pageSize
      )
    )

  /*
    Upload logic
    ------------
    Normal upload:
      replace only the transaction dates present in the file.

    Consolidated upload:
      replace the complete date range represented by the existing dataset
      plus the new consolidated file. This prevents old rows such as the
      previous 3,271-record snapshot from remaining behind when the new
      consolidated snapshot contains 2,438 rows.

    We deliberately use the already-loaded transaction dates rather than
    another client-side verification SELECT, because the RPC itself performs
    the authoritative inserted-row safety check.
  */
  async function uploadFile(e) {
    const file =
      e.target.files?.[0]

    if (!file) return

    setUploading(true)
    setError('')
    setMessage(
      'Reading and analysing Excel file...'
    )

    try {
      const buf =
        await file.arrayBuffer()

      const wb = XLSX.read(
        buf,
        {
          type: 'array',
          cellDates: false,
          raw: true
        }
      )

      const ws =
        wb.Sheets[
          wb.SheetNames[0]
        ]

      if (!ws) {
        throw new Error(
          'The Excel file does not contain a readable first sheet.'
        )
      }

      /*
        raw:true preserves Excel serial numbers. cellDates:false prevents
        SheetJS from creating Date objects that can introduce timezone shifts.
      */
      const raw =
        XLSX.utils.sheet_to_json(
          ws,
          {
            defval: null,
            raw: true
          }
        )

      if (!raw.length) {
        throw new Error(
          'No transaction rows were found in the Excel file.'
        )
      }

      const mapped =
        raw.map(mapRow)

      if (!mapped.length) {
        throw new Error(
          'No valid transactions found. Please use the normal Snowball transaction Excel format.'
        )
      }

      const invalidRows =
        mapped.filter(row =>
          !row.investor_name ||
          !row.transaction_date ||
          row.amount == null ||
          !Number.isFinite(
            Number(row.amount)
          )
        )

      if (invalidRows.length) {
        throw new Error(
          `Excel upload stopped: ${invalidRows.length} row(s) have missing Investor, Date, or Amount. No database data was changed.`
        )
      }

      /*
        IMPORTANT:
        classifyRows() is run on every row before upload.
        The database therefore receives the final classification, and the
        dashboard simply displays that stored classification.
      */
      const analysed =
        classifyRows(mapped).map(
          ({
            display_classification,
            ...row
          }) => ({
            ...row,
            classified_transaction_type:
              display_classification ||
              'Redemption',
            classification_status:
              'Completed',
            classification_reason:
              null
          })
        )

      const uploadDates = [
        ...new Set(
          analysed
            .map(
              x =>
                x.transaction_date
            )
            .filter(Boolean)
        )
      ]

      if (!uploadDates.length) {
        throw new Error(
          'No valid transaction dates were found in the Excel file.'
        )
      }

      const normalizedName =
        String(file.name || '')
          .trim()
          .toLowerCase()

      const isConsolidatedFile =
        normalizedName.includes('conso') ||
        normalizedName.includes(
          'consolidated'
        ) ||
        analysed.length >= 500

      /*
        For the consolidated snapshot, include every date currently in the
        application dataset. The SQL RPC deletes those dates first and then
        inserts every row from the new Excel snapshot.
      */
      const existingDates =
        isConsolidatedFile
          ? [
              ...new Set(
                rows
                  .map(
                    x =>
                      x.transaction_date
                  )
                  .filter(Boolean)
              )
            ]
          : []

      const datesToReplace =
        isConsolidatedFile
          ? [
              ...new Set([
                ...existingDates,
                ...uploadDates
              ])
            ]
          : uploadDates

      setMessage(
        `${isConsolidatedFile ? 'Replacing consolidated data' : 'Updating transactions'}: ${analysed.length} transactions across ${datesToReplace.length} date(s)...`
      )

      const {
        data,
        error: rpcError
      } =
        await supabase.rpc(
          'replace_transactions_for_dates',
          {
            p_dates:
              datesToReplace,
            p_rows:
              analysed
          }
        )

      if (rpcError) {
        throw rpcError
      }

      const inserted =
        Number(
          data?.inserted_transactions ??
          0
        )

      if (
        inserted !==
        analysed.length
      ) {
        throw new Error(
          `Upload verification failed. Excel contains ${analysed.length} transactions but Supabase inserted ${inserted}. No partial upload should be accepted.`
        )
      }

      await loadData()
      await loadRms()

      await notify(
        'Upload completed',
        `${inserted} transactions updated across ${uploadDates.length} Excel date(s).`
      )
    } catch (err) {
      console.error(err)

      setError(
        err?.message ||
        'Upload failed. Please try again.'
      )

      setMessage('')
    } finally {
      setUploading(false)
      e.target.value = ''
    }
  }

  function exportExcel() {
    const out =
      filtered.map(x => ({
        Date:
          x.transaction_date,
        RM:
          x.rm_name,
        Group:
          x.group_name,
        Investor:
          x.investor_name,
        Folio:
          x.folio_no,
        Scheme:
          x.scheme,
        Amount:
          x.amount,
        Source:
          sourceLabel(
            x.original_transaction_type
          ),
        Classification:
          x.display_classification
      }))

    const ws =
      XLSX.utils.json_to_sheet(
        out
      )

    ws['!cols'] = [
      { wch: 14 },
      { wch: 22 },
      { wch: 32 },
      { wch: 30 },
      { wch: 20 },
      { wch: 48 },
      { wch: 16 },
      { wch: 15 },
      { wch: 18 }
    ]

    const wb =
      XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      wb,
      ws,
      'Snowball Tracker'
    )

    XLSX.writeFile(
      wb,
      'snowball-transaction-report.xlsx'
    )
  }

  function exportPDF() {
    const doc =
      new jsPDF({
        orientation:
          'landscape',
        unit: 'mm',
        format: 'a4'
      })

    const pageWidth =
      doc.internal.pageSize.getWidth()

    const pageHeight =
      doc.internal.pageSize.getHeight()

    const pdfMoney = value =>
      `Rs. ${Number(
        value || 0
      ).toLocaleString(
        'en-IN',
        {
          maximumFractionDigits: 0
        }
      )}`

    doc.setFontSize(18)
    doc.setTextColor(
      36,
      55,
      75
    )
    doc.text(
      'Snowball Financial Services',
      12,
      14
    )

    doc.setFontSize(10.5)
    doc.setTextColor(
      90,
      105,
      120
    )
    doc.text(
      'Transaction & Redemption Analysis',
      12,
      21
    )

    doc.setFontSize(8)
    doc.text(
      `Generated: ${new Date().toLocaleString('en-IN')}`,
      pageWidth - 12,
      14,
      { align: 'right' }
    )

    autoTable(doc, {
      startY: 27,
      margin: {
        left: 12,
        right: 12,
        bottom: 14
      },

      head: [[
        'Date',
        'RM',
        'Group',
        'Investor',
        'Scheme',
        'Amount',
        'Source',
        'Classification'
      ]],

      body:
        filtered.map(x => [
          x.transaction_date ||
            '',
          x.rm_name || '',
          x.group_name || '',
          x.investor_name || '',
          x.scheme || '',
          pdfMoney(
            x.amount
          ),
          sourceLabel(
            x.original_transaction_type
          ),
          x.display_classification ||
            ''
        ]),

      theme: 'grid',

      styles: {
        fontSize: 8.2,
        cellPadding: 2.1,
        valign: 'middle',
        lineColor: [
          220,
          225,
          230
        ],
        lineWidth: 0.15,
        overflow: 'linebreak'
      },

      headStyles: {
        fillColor: [
          55,
          115,
          155
        ],
        textColor: [
          255,
          255,
          255
        ],
        fontStyle: 'bold',
        fontSize: 8.2,
        cellPadding: 2.2,
        halign: 'left'
      },

      alternateRowStyles: {
        fillColor: [
          247,
          249,
          251
        ]
      },

      columnStyles: {
        0: { cellWidth: 19 },
        1: { cellWidth: 24 },
        2: { cellWidth: 31 },
        3: { cellWidth: 34 },
        4: { cellWidth: 58 },
        5: {
          cellWidth: 25,
          halign: 'right'
        },
        6: {
          cellWidth: 24,
          halign: 'center'
        },
        7: {
          cellWidth: 30,
          halign: 'center'
        }
      },

      didDrawPage: function () {
        const pageNumber =
          doc.internal
            .getCurrentPageInfo()
            .pageNumber

        doc.setFontSize(8)
        doc.setTextColor(
          120,
          130,
          140
        )

        doc.text(
          `Page ${pageNumber}`,
          pageWidth - 12,
          pageHeight - 7,
          { align: 'right' }
        )
      }
    })

    doc.save(
      'snowball-transaction-report.pdf'
    )
  }

  if (!session) {
    if (forgotMode) {
      return (
        <main className="loginPage">
          <section className="loginCard">
            <button
              className="backLink"
              onClick={() => {
                setForgotMode(false)
                setError('')
                setForgotMessage('')
              }}
            >
              <ArrowLeft size={16} />
              Back to login
            </button>

            <h2>Reset Password</h2>

            <p className="loginSub">
              Enter your registered email address and we will send reset instructions.
            </p>

            <form
              onSubmit={
                sendResetPassword
              }
            >
              <label>
                Email Address
              </label>

              <input
                type="email"
                placeholder="Enter your email"
                value={
                  forgotEmail
                }
                onChange={e =>
                  setForgotEmail(
                    e.target.value
                  )
                }
                required
              />

              <button
                className="primaryButton fullButton"
                disabled={loading}
              >
                {loading
                  ? 'Sending...'
                  : 'Send Reset Instructions'}
              </button>

              {error && (
                <div className="error">
                  {error}
                </div>
              )}

              {forgotMessage && (
                <div className="successBox">
                  {
                    forgotMessage
                  }
                </div>
              )}
            </form>
          </section>
        </main>
      )
    }

    return (
      <main className="loginPage">
        <section className="loginCard">
          <h1>
            Snowball Redemption Tracker
          </h1>

          <p className="loginSub">
            Login to access the transaction dashboard.
          </p>

          <form onSubmit={login}>
            <label>
              Email Address
            </label>

            <input
              type="email"
              placeholder="Enter your email"
              value={email}
              onChange={e =>
                setEmail(
                  e.target.value
                )
              }
              required
            />

            <label>
              Password
            </label>

            <div className="passwordWrap">
              <input
                type={
                  showPassword
                    ? 'text'
                    : 'password'
                }
                placeholder="Enter your password"
                value={password}
                onChange={e =>
                  setPassword(
                    e.target.value
                  )
                }
                required
              />

              <button
                type="button"
                className="passwordToggle"
                onClick={() =>
                  setShowPassword(
                    !showPassword
                  )
                }
              >
                {showPassword
                  ? <EyeOff size={18} />
                  : <Eye size={18} />}
              </button>
            </div>

            <button
              className="primaryButton fullButton"
              disabled={loading}
            >
              {loading
                ? 'Logging in...'
                : 'Login'}
            </button>

            <button
              type="button"
              className="forgotLink"
              onClick={() => {
                setForgotMode(true)
                setError('')
              }}
            >
              Forgot / Reset Password?
            </button>

            {error && (
              <div className="error">
                {error}
              </div>
            )}
          </form>
        </section>
      </main>
    )
  }

  const title =
    page === 'dashboard'
      ? 'Snowball Redemption Tracker'
      : page === 'transactions'
        ? 'Transaction Data'
        : 'Settings'

  return (
    <div className="appShell">
      <aside className="sidebar">
        <div className="sidebarBrand">
          <img
            src={logoUrl}
            alt="Snowball Financial Services"
          />
        </div>

        <nav className="sidebarNav">
          <button
            className={
              page === 'dashboard'
                ? 'navActive'
                : ''
            }
            onClick={() =>
              setPage('dashboard')
            }
          >
            <LayoutDashboard
              size={18}
            />
            Dashboard
          </button>

          <button
            className={
              page === 'transactions'
                ? 'navActive'
                : ''
            }
            onClick={() =>
              setPage(
                'transactions'
              )
            }
          >
            <Table2 size={18} />
            Transaction Data
          </button>

          <button
            className={
              page === 'settings'
                ? 'navActive'
                : ''
            }
            onClick={() =>
              setPage('settings')
            }
          >
            <Settings
              size={18}
            />
            Settings
          </button>
        </nav>

        <div className="sidebarBottom">
          <button
            onClick={
              handleLogout
            }
          >
            <LogOut size={18} />
            Logout
          </button>
        </div>
      </aside>

      <main className="mainContent">
        <header className="topHeader">
          <div>
            <h1>{title}</h1>

            <p>
              {page ===
                'dashboard' &&
                'Analyse transactions and monitor redemption activity'}

              {page ===
                'transactions' &&
                'Detailed transaction-level data and classifications'}

              {page ===
                'settings' &&
                'Manage Relationship Managers and administrator access'}
            </p>
          </div>

          <div className="headerActions">
            <div className="notificationWrap">
              <button
                className="iconButton notificationBell"
                onClick={async () => {
                  setShowNotifications(
                    v => !v
                  )

                  if (
                    !showNotifications
                  ) {
                    await markNotificationsRead()
                  }
                }}
                title="Notifications"
              >
                <Bell size={18} />

                {notifications.filter(
                  x => !x.is_read
                ).length > 0 && (
                  <span className="notificationCount">
                    {
                      notifications.filter(
                        x => !x.is_read
                      ).length
                    }
                  </span>
                )}
              </button>

              {showNotifications && (
                <div className="notificationPanel">
                  <div className="notificationPanelHead">
                    <strong>
                      Notifications
                    </strong>

                    <button
                      onClick={() =>
                        setShowNotifications(
                          false
                        )
                      }
                    >
                      <X size={16} />
                    </button>
                  </div>

                  {notifications.length === 0
                    ? (
                      <p className="notificationEmpty">
                        No notifications yet.
                      </p>
                    )
                    : notifications.map(
                        n => (
                          <div
                            className={`notificationItem ${
                              n.is_read
                                ? ''
                                : 'unread'
                            }`}
                            key={n.id}
                          >
                            <strong>
                              {n.title}
                            </strong>

                            <span>
                              {n.body}
                            </span>

                            <small>
                              {new Date(
                                n.created_at
                              ).toLocaleString(
                                'en-IN'
                              )}
                            </small>
                          </div>
                        )
                      )}
                </div>
              )}
            </div>

            <button
              className="iconButton"
              onClick={async () => {
                await loadData()
                await loadRms()
                await loadNotifications()
              }}
              title="Refresh"
            >
              <RefreshCw
                size={18}
              />
            </button>

            <div
              className="userAvatar"
              title={
                session.user.email
              }
            >
              {session.user.email
                ?.charAt(0)
                .toUpperCase()}
            </div>
          </div>
        </header>

        {error && (
          <div className="error">
            {error}
          </div>
        )}

        {message && (
          <div className="message">
            {message}
          </div>
        )}

        {page !== 'settings' && (
          <section className="filterBar">
            <div className="periodButtons">
              {[
                'WTD',
                'MTD',
                'QTD',
                'YTD'
              ].map(p => (
                <button
                  key={p}
                  className={
                    period === p &&
                    !from &&
                    !to
                      ? 'activePeriod'
                      : ''
                  }
                  onClick={() => {
                    setPeriod(p)
                    setFrom('')
                    setTo('')
                  }}
                >
                  {p}
                </button>
              ))}
            </div>

            <select
              value={rm}
              onChange={e =>
                setRm(
                  e.target.value
                )
              }
            >
              {activeRmNames.map(
                x => (
                  <option
                    key={x}
                    value={x}
                  >
                    {x}
                  </option>
                )
              )}
            </select>

            <input
              type="date"
              value={from}
              onChange={e =>
                setFrom(
                  e.target.value
                )
              }
            />

            <input
              type="date"
              value={to}
              onChange={e =>
                setTo(
                  e.target.value
                )
              }
            />

            {page ===
              'transactions' && (
              <>
                <button className="uploadButton">
                  <Upload
                    size={17}
                  />

                  <label>
                    {uploading
                      ? 'Uploading...'
                      : 'Upload Excel'}

                    <input
                      type="file"
                      accept=".xlsx,.xls,.csv"
                      onChange={
                        uploadFile
                      }
                      disabled={
                        uploading
                      }
                    />
                  </label>
                </button>

                <button
                  className="secondaryButton"
                  onClick={
                    exportExcel
                  }
                >
                  <Download
                    size={16}
                  />
                  Excel
                </button>

                <button
                  className="secondaryButton"
                  onClick={
                    exportPDF
                  }
                >
                  <FileText
                    size={16}
                  />
                  PDF
                </button>
              </>
            )}
          </section>
        )}

        {page ===
          'dashboard' && (
          <>
            <section className="kpiGrid">
              {[
                [
                  'Redemption',
                  totals.Redemption,
                  'redemption'
                ],
                [
                  'SWP',
                  totals.SWP,
                  'swp'
                ],
                [
                  'Switch',
                  totals.Switch,
                  'switch'
                ],
                [
                  'STP',
                  totals.STP,
                  'stp'
                ],
                [
                  'Investors',
                  totals.Investors,
                  'investors'
                ],
                [
                  'Transactions',
                  totals.Transactions,
                  'transactions'
                ]
              ].map(
                ([
                  name,
                  value,
                  tone
                ]) => (
                  <article
                    className={`kpiCard ${tone}`}
                    key={name}
                  >
                    <span>
                      {name}
                    </span>

                    <strong>
                      {[
                        'Investors',
                        'Transactions'
                      ].includes(
                        name
                      )
                        ? Number(
                            value
                          ).toLocaleString(
                            'en-IN'
                          )
                        : money(
                            value
                          )}
                    </strong>
                  </article>
                )
              )}
            </section>

            <section className="dashboardGrid">
              <article className="chartCard classificationCard">
                <h3>
                  Amount by Classification
                </h3>

                <div className="donutLayout">
                  <div
                    className="donut"
                    style={
                      donutStyle
                    }
                  >
                    <div className="donutHole">
                      <strong>
                        {totals.Transactions.toLocaleString(
                          'en-IN'
                        )}
                      </strong>

                      <span>
                        Transactions
                      </span>
                    </div>
                  </div>

                  <div className="classificationLegend">
                    {classificationChart.map(
                      item => (
                        <div
                          className="legendRow"
                          key={
                            item.name
                          }
                        >
                          <span
                            className="legendDot"
                            style={{
                              background:
                                item.color
                            }}
                          />

                          <span>
                            {
                              item.name
                            }
                          </span>

                          <strong>
                            {compactMoney(
                              item.amount
                            )}
                          </strong>
                        </div>
                      )
                    )}
                  </div>
                </div>
              </article>

              <article className="chartCard">
                <h3>
                  Monthly Trend
                </h3>

                <div className="trendChart">
                  {monthlyTrend.length
                    ? monthlyTrend.map(
                        item => {
                          const max =
                            Math.max(
                              ...monthlyTrend.flatMap(
                                x => [
                                  x.Redemption,
                                  x.SWP,
                                  x.Switch,
                                  x.STP
                                ]
                              ),
                              1
                            )

                          return (
                            <div
                              className="trendMonth"
                              key={
                                item.month
                              }
                            >
                              <div className="trendBars">
                                {[
                                  [
                                    'Redemption',
                                    'trendRed'
                                  ],
                                  [
                                    'SWP',
                                    'trendSwp'
                                  ],
                                  [
                                    'Switch',
                                    'trendSwitch'
                                  ],
                                  [
                                    'STP',
                                    'trendStp'
                                  ]
                                ].map(
                                  ([
                                    key,
                                    cls
                                  ]) => (
                                    <div
                                      key={
                                        key
                                      }
                                      className={`trendBar ${cls}`}
                                      title={`${key}: ${money(item[key])}`}
                                      style={{
                                        height:
                                          `${Math.max(
                                            4,
                                            (
                                              item[key] /
                                              max
                                            ) * 100
                                          )}%`
                                      }}
                                    />
                                  )
                                )}
                              </div>

                              <span>
                                {item.month.slice(
                                  5
                                )}
                              </span>
                            </div>
                          )
                        }
                      )
                    : (
                      <div className="emptyChart">
                        No data available
                      </div>
                    )}
                </div>
              </article>

              <article className="chartCard rmChartCard">
                <h3>
                  Transactions by RM
                </h3>

                <div className="rmBars">
                  {chartData.map(
                    item => (
                      <div
                        className="rmBarRow"
                        key={
                          item.name
                        }
                      >
                        <span
                          title={
                            item.name
                          }
                        >
                          {item.name}
                        </span>

                        <div className="rmBarTrack">
                          <div
                            className="rmBarFill"
                            style={{
                              width:
                                `${Math.max(
                                  3,
                                  (
                                    item.amount /
                                    maxChartValue
                                  ) * 100
                                )}%`
                            }}
                          />
                        </div>

                        <strong>
                          {compactMoney(
                            item.amount
                          )}
                        </strong>
                      </div>
                    )
                  )}

                  {!chartData.length && (
                    <div className="emptyChart">
                      No data available
                    </div>
                  )}
                </div>
              </article>
            </section>

            <section className="recentCard">
              <div className="sectionHeading">
                <div>
                  <h2>
                    Recent Transactions
                  </h2>

                  <p>
                    Latest analysed transactions
                  </p>
                </div>

                <button
                  className="viewAllButton"
                  onClick={() =>
                    setPage(
                      'transactions'
                    )
                  }
                >
                  View All Transactions
                </button>
              </div>

              <div className="miniTableWrap">
                <table>
                  <thead>
                    <tr>
                      <th>Date</th>
                      <th>RM</th>
                      <th>Investor</th>
                      <th>Scheme</th>
                      <th>Amount</th>
                      <th>Source</th>
                      <th>Classification</th>
                    </tr>
                  </thead>

                  <tbody>
                    {filtered
                      .slice(0, 10)
                      .map(x => (
                        <tr
                          key={
                            x.id ||
                            `${x.investor_name}-${x.transaction_date}-${x.scheme}-${x.amount}`
                          }
                        >
                          <td>
                            {
                              x.transaction_date
                            }
                          </td>
                          <td>
                            {
                              x.rm_name
                            }
                          </td>
                          <td>
                            {
                              x.investor_name
                            }
                          </td>
                          <td>
                            {
                              x.scheme
                            }
                          </td>
                          <td>
                            {money(
                              x.amount
                            )}
                          </td>
                          <td>
                            {sourceLabel(
                              x.original_transaction_type
                            )}
                          </td>
                          <td>
                            <span
                              className={`classification ${
                                String(
                                  x.display_classification
                                ).toLowerCase()
                              }`}
                            >
                              {
                                x.display_classification
                              }
                            </span>
                          </td>
                        </tr>
                      ))}
                  </tbody>
                </table>
              </div>
            </section>
          </>
        )}

        {page ===
          'transactions' && (
          <section className="dataCard">
            <div className="sectionHeading">
              <div>
                <h2>
                  Transaction Details
                </h2>

                <p>
                  {filtered.length}{' '}
                  transactions found
                </p>
              </div>
            </div>

            <div className="tableScroll">
              <table>
                <thead>
                  <tr>
                    <th>Date</th>
                    <th>RM</th>
                    <th>Investor</th>
                    <th>Folio</th>
                    <th>Scheme</th>
                    <th>Amount</th>
                    <th>Source</th>
                    <th>Classification</th>
                  </tr>
                </thead>

                <tbody>
                  {paginatedRows.map(
                    x => (
                      <tr
                        key={
                          x.id ||
                          `${x.investor_name}-${x.transaction_date}-${x.scheme}-${x.amount}-${x.rm_name}`
                        }
                      >
                        <td>
                          {
                            x.transaction_date
                          }
                        </td>
                        <td>
                          {
                            x.rm_name
                          }
                        </td>
                        <td>
                          {
                            x.investor_name
                          }
                        </td>
                        <td>
                          {
                            x.folio_no
                          }
                        </td>
                        <td>
                          {
                            x.scheme
                          }
                        </td>
                        <td>
                          {money(
                            x.amount
                          )}
                        </td>
                        <td>
                          {sourceLabel(
                            x.original_transaction_type
                          )}
                        </td>
                        <td>
                          <span
                            className={`classification ${
                              String(
                                x.display_classification
                              ).toLowerCase()
                            }`}
                          >
                            {
                              x.display_classification
                            }
                          </span>
                        </td>
                      </tr>
                    )
                  )}
                </tbody>
              </table>
            </div>

            <div className="pagination">
              <button
                disabled={
                  currentPage === 1
                }
                onClick={() =>
                  setCurrentPage(
                    currentPage - 1
                  )
                }
              >
                Previous
              </button>

              <span>
                Page {currentPage} of{' '}
                {totalPages}
              </span>

              <button
                disabled={
                  currentPage ===
                  totalPages
                }
                onClick={() =>
                  setCurrentPage(
                    currentPage + 1
                  )
                }
              >
                Next
              </button>
            </div>
          </section>
        )}

        {page ===
          'settings' && (
          <section className="settingsPage">
            <div className="settingsTabs">
              <button
                className={
                  settingsTab ===
                  'current'
                    ? 'tabActive'
                    : ''
                }
                onClick={() =>
                  setSettingsTab(
                    'current'
                  )
                }
              >
                <Users size={16} />
                Current RMs
              </button>

              <button
                className={
                  settingsTab ===
                  'admin'
                    ? 'tabActive'
                    : ''
                }
                onClick={() =>
                  setSettingsTab(
                    'admin'
                  )
                }
              >
                <ShieldCheck
                  size={16}
                />
                Admin Settings
              </button>
            </div>

            {settingsTab ===
              'current' && (
              <>
                <article className="rmAddCard">
                  <h2>
                    Add New RM
                  </h2>

                  <p>
                    Add an RM to the organisation. Existing RMs are never deleted; they can be marked inactive.
                  </p>

                  <form
                    onSubmit={addRm}
                  >
                    <input
                      placeholder="Enter RM name"
                      value={newRm}
                      onChange={e =>
                        setNewRm(
                          e.target.value
                        )
                      }
                    />

                    <button className="primaryButton">
                      <UserPlus
                        size={16}
                      />
                      Add RM
                    </button>
                  </form>
                </article>

                <article className="rmListCard">
                  <div className="sectionHeading">
                    <div>
                      <h2>
                        Current RMs
                      </h2>

                      <p>
                        {
                          rmDirectory.filter(
                            x =>
                              x.is_active !== false
                          ).length
                        } active
                        {' • '}
                        {
                          rmDirectory.length
                        } total
                      </p>
                    </div>
                  </div>

                  <div className="rmList">
                    {rmDirectory.map(
                      item => (
                        <div
                          className="rmRow"
                          key={
                            item.id ||
                            item.name
                          }
                        >
                          <div className="rmAvatar">
                            {item.name
                              ?.charAt(
                                0
                              )
                              .toUpperCase()}
                          </div>

                          <div className="rmName">
                            <strong>
                              {
                                item.name
                              }
                            </strong>

                            <span
                              className={
                                item.is_active === false
                                  ? 'statusInactive'
                                  : 'statusActive'
                              }
                            >
                              {item.is_active === false
                                ? 'Inactive'
                                : 'Active'}
                            </span>
                          </div>

                          <button
                            className={
                              item.is_active === false
                                ? 'activateButton'
                                : 'deactivateButton'
                            }
                            onClick={() =>
                              toggleRmStatus(
                                item
                              )
                            }
                          >
                            {item.is_active === false
                              ? (
                                <>
                                  <UserPlus
                                    size={14}
                                  />
                                  Activate
                                </>
                              )
                              : (
                                <>
                                  <UserMinus
                                    size={14}
                                  />
                                  Mark Inactive
                                </>
                              )}
                          </button>
                        </div>
                      )
                    )}
                  </div>
                </article>
              </>
            )}

            {settingsTab ===
              'admin' && (
              <article className="rmListCard">
                <div className="sectionHeading">
                  <div>
                    <h2>
                      Admin Settings
                    </h2>

                    <p>
                      Select which active Relationship Managers should have administrator access.
                    </p>
                  </div>
                </div>

                <div className="rmList">
                  {rmDirectory.map(
                    item => (
                      <div
                        className="rmRow"
                        key={
                          item.id ||
                          item.name
                        }
                      >
                        <div className="rmAvatar">
                          {item.name
                            ?.charAt(
                              0
                            )
                            .toUpperCase()}
                        </div>

                        <div className="rmName">
                          <strong>
                            {
                              item.name
                            }
                          </strong>

                          {item.is_admin &&
                            item.admin_email && (
                              <small className="adminEmail">
                                {
                                  item.admin_email
                                }
                              </small>
                            )}

                          <span
                            className={
                              item.is_admin
                                ? 'statusAdmin'
                                : 'statusInactive'
                            }
                          >
                            {item.is_admin
                              ? 'Admin'
                              : 'Standard access'}
                          </span>
                        </div>

                        <button
                          className={
                            item.is_admin
                              ? 'deactivateButton'
                              : 'activateButton'
                          }
                          onClick={() =>
                            openAdminModal(
                              item
                            )
                          }
                        >
                          {item.is_admin
                            ? (
                              <>
                                <CheckCircle2
                                  size={14}
                                />
                                Remove Admin
                              </>
                            )
                            : (
                              <>
                                <ShieldCheck
                                  size={14}
                                />
                                Make Admin
                              </>
                            )}
                        </button>
                      </div>
                    )
                  )}
                </div>
              </article>
            )}
          </section>
        )}

        {adminModal && (
          <div
            className="modalBackdrop"
            onMouseDown={() =>
              setAdminModal(null)
            }
          >
            <div
              className="adminModal"
              onMouseDown={e =>
                e.stopPropagation()
              }
            >
              <div className="modalHeader">
                <div>
                  <h2>
                    Grant Admin Access
                  </h2>

                  <p>
                    Enter the email address this employee will use to log in.
                  </p>
                </div>

                <button
                  className="modalClose"
                  onClick={() =>
                    setAdminModal(
                      null
                    )
                  }
                >
                  <X size={18} />
                </button>
              </div>

              <form
                onSubmit={
                  confirmAdmin
                }
              >
                <label>
                  Login Email Address
                </label>

                <div className="emailField">
                  <Mail
                    size={16}
                  />

                  <input
                    autoFocus
                    type="email"
                    placeholder="name@snowballwealth.in"
                    value={
                      adminEmail
                    }
                    onChange={e =>
                      setAdminEmail(
                        e.target.value
                      )
                    }
                  />
                </div>

                <p className="modalNote">
                  Admin access is linked to this email. The email must also have a valid Supabase Auth account to sign in.
                </p>

                <div className="modalActions">
                  <button
                    type="button"
                    onClick={() =>
                      setAdminModal(
                        null
                      )
                    }
                  >
                    Cancel
                  </button>

                  <button
                    className="primaryButton"
                    type="submit"
                  >
                    <ShieldCheck
                      size={16}
                    />
                    Grant Admin Access
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}
      </main>
    </div>
  )
}

createRoot(
  document.getElementById('root')
).render(<App />)
