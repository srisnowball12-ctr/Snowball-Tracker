import React, { useEffect, useMemo, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'
import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'

import {
  LayoutDashboard,
  Table2,
  Users,
  Settings as SettingsIcon,
  LogOut,
  RefreshCw,
  Upload,
  Download,
  FileText,
  UserPlus,
  UserMinus,
  UserCheck,
  Eye,
  EyeOff,
  ArrowLeft,
  Mail,
  Shield,
  Plus,
  X
} from 'lucide-react'

import './styles.css'
import logo from './logo.png'

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
)

const money = value =>
  new Intl.NumberFormat('en-IN', {
    style: 'currency',
    currency: 'INR',
    maximumFractionDigits: 0
  }).format(Number(value || 0))

const compactMoney = value => {
  const n = Number(value || 0)

  if (n >= 10000000) {
    return `₹${(n / 10000000).toFixed(2)} Cr`
  }

  if (n >= 100000) {
    return `₹${(n / 100000).toFixed(2)} L`
  }

  if (n >= 1000) {
    return `₹${(n / 1000).toFixed(1)} K`
  }

  return money(n)
}

const norm = value =>
  String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')

const iso = value => {
  if (!value) return null

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10)
  }

  const date = new Date(value)

  if (Number.isNaN(date.getTime())) {
    return null
  }

  return date.toISOString().slice(0, 10)
}

const normalizeSource = value => {
  const text = String(value || '')
    .trim()
    .toLowerCase()

  if (text.includes('swp')) return 'SWP'
  if (text.includes('stp')) return 'STP'
  if (text.includes('switch')) return 'Switch'

  if (
    text.includes('redemption') ||
    text === 'red' ||
    text.includes('redeem')
  ) {
    return 'Redemption'
  }

  return 'Redemption'
}

function mapRow(row) {
  const lookup = Object.fromEntries(
    Object.entries(row).map(([key, value]) => [
      norm(key),
      value
    ])
  )

  const get = (...keys) =>
    keys
      .map(key => lookup[norm(key)])
      .find(value => value !== undefined)

  const amountRaw = get(
    'Amount(₹)',
    'Amount',
    'Transaction Amount',
    'amount'
  )

  const amount =
    typeof amountRaw === 'number'
      ? amountRaw
      : Number(
          String(amountRaw || '')
            .replace(/[₹,\s]/g, '')
            .replace(/,/g, '')
        )

  const originalType =
    get(
      'Type',
      'Transaction Type',
      'original_transaction_type'
    ) || 'Redemption'

  return {
    rm_name:
      get(
        'Partner/Employee',
        'RM',
        'Relationship Manager',
        'rm_name'
      ) || null,

    group_name:
      get('Group', 'group_name') || null,

    investor_name:
      get(
        'Investor',
        'Investor Name',
        'Client Name',
        'investor_name'
      ) || null,

    transaction_date: iso(
      get(
        'Date',
        'Transaction Date',
        'transaction_date'
      )
    ),

    folio_no:
      String(
        get(
          'Folio No/Demat A/C',
          'Folio No',
          'Folio',
          'folio_no'
        ) || ''
      ) || null,

    scheme:
      get('Scheme', 'Fund', 'scheme') || null,

    amount: Number.isFinite(amount)
      ? amount
      : null,

    transaction_type: originalType,

    original_transaction_type: originalType,

    classified_transaction_type: null,

    classification_status: 'Completed',

    classification_reason: null
  }
}

/*
  Classification logic

  1. If Excel source clearly says SWP/STP/Switch,
     that classification is retained.

  2. If source says Redemption/Red,
     system checks recurring investor + scheme history.

  3. If at least 3 transactions show approximately
     monthly recurring dates and similar amounts,
     it is classified as SWP.
*/

function daysBetween(a, b) {
  const first = new Date(a + 'T00:00:00')
  const second = new Date(b + 'T00:00:00')

  return Math.abs(
    Math.round(
      (second - first) /
      (1000 * 60 * 60 * 24)
    )
  )
}

function hasRecurringSwpPattern(row, allRows) {
  if (
    !row.investor_name ||
    !row.scheme ||
    !row.transaction_date
  ) {
    return false
  }

  const matches = allRows
    .filter(item =>
      norm(item.investor_name) === norm(row.investor_name) &&
      norm(item.scheme) === norm(row.scheme)
    )
    .filter(item => item.transaction_date)
    .sort((a, b) =>
      String(a.transaction_date).localeCompare(
        String(b.transaction_date)
      )
    )

  if (matches.length < 3) {
    return false
  }

  const amounts = matches
    .map(item => Number(item.amount || 0))
    .filter(value => value > 0)

  if (amounts.length < 3) {
    return false
  }

  const average =
    amounts.reduce((sum, value) => sum + value, 0) /
    amounts.length

  const amountVariation =
    Math.max(...amounts) -
    Math.min(...amounts)

  const similarAmounts =
    average > 0 &&
    amountVariation / average <= 0.30

  if (!similarAmounts) {
    return false
  }

  const intervals = []

  for (let i = 1; i < matches.length; i++) {
    intervals.push(
      daysBetween(
        matches[i].transaction_date,
        matches[i - 1].transaction_date
      )
    )
  }

  const monthlyIntervals = intervals.filter(
    days => days >= 20 && days <= 40
  ).length

  return monthlyIntervals >= 2
}

function getClassification(row, allRows) {
  const source = normalizeSource(
    row.original_transaction_type ||
    row.transaction_type
  )

  if (source === 'SWP') return 'SWP'
  if (source === 'STP') return 'STP'
  if (source === 'Switch') return 'Switch'

  if (hasRecurringSwpPattern(row, allRows)) {
    return 'SWP'
  }

  return 'Redemption'
}

function App() {
  const [session, setSession] = useState(null)

  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [showPassword, setShowPassword] =
    useState(false)

  const [forgotMode, setForgotMode] =
    useState(false)

  const [forgotEmail, setForgotEmail] =
    useState('')

  const [forgotMessage, setForgotMessage] =
    useState('')

  const [page, setPage] =
    useState('dashboard')

  const [rows, setRows] =
    useState([])

  const [rms, setRms] =
    useState([])

  const [adminEmails, setAdminEmails] =
    useState([])

  const [isAdmin, setIsAdmin] =
    useState(false)

  const [rm, setRm] =
    useState('All')

  const [from, setFrom] =
    useState('')

  const [to, setTo] =
    useState('')

  const [period, setPeriod] =
    useState('YTD')

  const [loading, setLoading] =
    useState(false)

  const [uploading, setUploading] =
    useState(false)

  const [error, setError] =
    useState('')

  const [message, setMessage] =
    useState('')

  const [newRm, setNewRm] =
    useState('')

  const [rmView, setRmView] =
    useState('current')

  const [newAdminEmail, setNewAdminEmail] =
    useState('')

  const [currentPage, setCurrentPage] =
    useState(1)

  const pageSize = 50

  useEffect(() => {
    supabase.auth
      .getSession()
      .then(({ data }) => {
        setSession(data.session)
      })

    const {
      data: { subscription }
    } = supabase.auth.onAuthStateChange(
      (_event, newSession) => {
        setSession(newSession)

        if (!newSession) {
          setPage('dashboard')
        }
      }
    )

    return () => subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (session) {
      loadData()
      loadRms()
      checkAdmin()
    }
  }, [session])

  useEffect(() => {
    setCurrentPage(1)
  }, [rm, from, to, period])

  async function login(event) {
    event.preventDefault()

    setLoading(true)
    setError('')

    const { error: loginError } =
      await supabase.auth.signInWithPassword({
        email,
        password
      })

    if (loginError) {
      setError(loginError.message)
    }

    setLoading(false)
  }

  async function sendResetPassword(event) {
    event.preventDefault()

    setLoading(true)
    setError('')
    setForgotMessage('')

    const { error: resetError } =
      await supabase.auth.resetPasswordForEmail(
        forgotEmail,
        {
          redirectTo: window.location.origin
        }
      )

    if (resetError) {
      setError(resetError.message)
    } else {
      setForgotMessage(
        'Password reset instructions have been sent.'
      )
    }

    setLoading(false)
  }

  async function logout() {
    setLoading(true)

    await supabase.auth.signOut()

    setSession(null)
    setRows([])
    setRms([])
    setAdminEmails([])
    setPage('dashboard')

    setLoading(false)
  }

  async function loadData() {
    const { data, error: dataError } =
      await supabase
        .from('transactions')
        .select('*')
        .order('transaction_date', {
          ascending: false
        })

    if (dataError) {
      setError(dataError.message)
      return
    }

    setRows(data || [])
  }

  async function loadRms() {
    const { data, error: rmError } =
      await supabase
        .from('rms')
        .select('*')
        .order('name')

    if (rmError) {
      setError(rmError.message)
      return
    }

    setRms(data || [])
  }

  async function checkAdmin() {
    if (!session?.user?.email) return

    const { data, error: adminError } =
      await supabase
        .from('app_admins')
        .select('*')
        .eq(
          'email',
          session.user.email
        )
        .eq('is_active', true)

    if (adminError) {
      console.error(adminError)
      setIsAdmin(false)
      return
    }

    setIsAdmin(Boolean(data?.length))
  }

  async function loadAdmins() {
    if (!isAdmin) return

    const { data, error: adminError } =
      await supabase
        .from('app_admins')
        .select('*')
        .order('email')

    if (adminError) {
      setError(adminError.message)
      return
    }

    setAdminEmails(data || [])
  }

  useEffect(() => {
    if (
      session &&
      isAdmin &&
      page === 'settings'
    ) {
      loadAdmins()
    }
  }, [session, isAdmin, page])

  async function addRm(event) {
    event.preventDefault()

    const name = newRm.trim()

    if (!name) return

    setError('')

    const { error: insertError } =
      await supabase
        .from('rms')
        .insert({
          name,
          is_active: true
        })

    if (insertError) {
      if (insertError.code === '23505') {
        setError('This RM already exists.')
      } else {
        setError(insertError.message)
      }

      return
    }

    setNewRm('')
    setMessage(`${name} added successfully.`)

    await loadRms()
  }

  async function toggleRmStatus(item) {
    setError('')

    const { error: updateError } =
      await supabase
        .from('rms')
        .update({
          is_active: !item.is_active
        })
        .eq('id', item.id)

    if (updateError) {
      setError(updateError.message)
      return
    }

    setMessage(
      `${item.name} is now ${
        item.is_active
          ? 'Inactive'
          : 'Active'
      }.`
    )

    await loadRms()
  }

  async function addAdmin(event) {
    event.preventDefault()

    const adminEmail =
      newAdminEmail.trim().toLowerCase()

    if (!adminEmail) return

    const { error: insertError } =
      await supabase
        .from('app_admins')
        .insert({
          email: adminEmail,
          is_active: true
        })

    if (insertError) {
      if (insertError.code === '23505') {
        setError(
          'This email is already in the Admin list.'
        )
      } else {
        setError(insertError.message)
      }

      return
    }

    setNewAdminEmail('')
    setMessage('Admin email added successfully.')

    await loadAdmins()
  }

  async function toggleAdminStatus(item) {
    const { error: updateError } =
      await supabase
        .from('app_admins')
        .update({
          is_active: !item.is_active
        })
        .eq('id', item.id)

    if (updateError) {
      setError(updateError.message)
      return
    }

    await loadAdmins()
  }

  const activeRmNames = useMemo(() => {
    return [
      'All',
      ...rms
        .filter(item => item.is_active)
        .map(item => item.name)
    ]
  }, [rms])

  /*
    Re-classify all rows for display.

    This means older transactions already stored
    in the database can also be analysed again.
  */

  const analysedRows = useMemo(() => {
    return rows.map(row => ({
      ...row,
      source_type: normalizeSource(
        row.original_transaction_type ||
        row.transaction_type
      ),
      system_classification:
        getClassification(row, rows)
    }))
  }, [rows])

  const filtered = useMemo(() => {
    return analysedRows.filter(item => {
      if (
        rm !== 'All' &&
        item.rm_name !== rm
      ) {
        return false
      }

      const date = item.transaction_date

      if (!date) return false

      if (from && date < from) {
        return false
      }

      if (to && date > to) {
        return false
      }

      if (!from && !to) {
        const now = new Date()

        const transactionDate =
          new Date(date + 'T00:00:00')

        if (period === 'WTD') {
          const day =
            (now.getDay() + 6) % 7

          const start =
            new Date(now)

          start.setDate(
            now.getDate() - day
          )

          start.setHours(
            0,
            0,
            0,
            0
          )

          if (transactionDate < start) {
            return false
          }
        }

        if (
          period === 'MTD' &&
          (
            transactionDate.getMonth() !==
              now.getMonth() ||
            transactionDate.getFullYear() !==
              now.getFullYear()
          )
        ) {
          return false
        }

        if (period === 'QTD') {
          const quarter =
            Math.floor(now.getMonth() / 3)

          if (
            transactionDate.getFullYear() !==
              now.getFullYear() ||
            Math.floor(
              transactionDate.getMonth() / 3
            ) !== quarter
          ) {
            return false
          }
        }

        if (
          period === 'YTD' &&
          transactionDate.getFullYear() !==
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
    const totalFor = classification =>
      filtered
        .filter(
          item =>
            item.system_classification ===
            classification
        )
        .reduce(
          (sum, item) =>
            sum + Number(item.amount || 0),
          0
        )

    return {
      Redemption: totalFor('Redemption'),
      SWP: totalFor('SWP'),
      Switch: totalFor('Switch'),
      STP: totalFor('STP'),

      Investors: new Set(
        filtered
          .map(item => item.investor_name)
          .filter(Boolean)
      ).size,

      Transactions: filtered.length
    }
  }, [filtered])

  const classificationData = useMemo(() => {
    return [
      {
        name: 'Redemption',
        amount: totals.Redemption
      },
      {
        name: 'SWP',
        amount: totals.SWP
      },
      {
        name: 'Switch',
        amount: totals.Switch
      },
      {
        name: 'STP',
        amount: totals.STP
      }
    ]
  }, [totals])

  const rmChartData = useMemo(() => {
    const values = {}

    filtered.forEach(item => {
      const name =
        item.rm_name || 'Unassigned'

      values[name] =
        (values[name] || 0) +
        Number(item.amount || 0)
    })

    return Object.entries(values)
      .map(([name, amount]) => ({
        name,
        amount
      }))
      .sort(
        (a, b) =>
          b.amount - a.amount
      )
      .slice(0, 7)
  }, [filtered])

  const monthlyTrend = useMemo(() => {
    const months = []

    for (let i = 5; i >= 0; i--) {
      const date = new Date()
      date.setDate(1)
      date.setMonth(date.getMonth() - i)

      const key =
        `${date.getFullYear()}-${String(
          date.getMonth() + 1
        ).padStart(2, '0')}`

      months.push({
        key,
        label: date.toLocaleString(
          'en-IN',
          {
            month: 'short'
          }
        ),
        Redemption: 0,
        SWP: 0,
        Switch: 0,
        STP: 0
      })
    }

    filtered.forEach(item => {
      if (!item.transaction_date) return

      const key =
        item.transaction_date.slice(0, 7)

      const month =
        months.find(
          value => value.key === key
        )

      if (
        month &&
        month[
          item.system_classification
        ] !== undefined
      ) {
        month[
          item.system_classification
        ] += Number(item.amount || 0)
      }
    })

    return months
  }, [filtered])

  const maxRmAmount = Math.max(
    ...rmChartData.map(
      item => item.amount
    ),
    1
  )

  const maxTrendAmount = Math.max(
    ...monthlyTrend.flatMap(item => [
      item.Redemption,
      item.SWP,
      item.Switch,
      item.STP
    ]),
    1
  )

  const donutTotal =
    classificationData.reduce(
      (sum, item) =>
        sum + item.amount,
      0
    ) || 1

  const redemptionPercent =
    totals.Redemption /
    donutTotal *
    100

  const swpPercent =
    totals.SWP /
    donutTotal *
    100

  const switchPercent =
    totals.Switch /
    donutTotal *
    100

  const stpPercent =
    totals.STP /
    donutTotal *
    100

  const donutStyle = {
    background: `conic-gradient(
      #d75b61 0 ${redemptionPercent}%,
      #5b987f ${redemptionPercent}% ${
        redemptionPercent + swpPercent
      }%,
      #6380a7 ${
        redemptionPercent + swpPercent
      }% ${
        redemptionPercent +
        swpPercent +
        switchPercent
      }%,
      #8c70ae ${
        redemptionPercent +
        swpPercent +
        switchPercent
      }% 100%
    )`
  }

  const totalPages = Math.max(
    1,
    Math.ceil(
      filtered.length / pageSize
    )
  )

  const paginatedRows = useMemo(() => {
    const start =
      (currentPage - 1) * pageSize

    return filtered.slice(
      start,
      start + pageSize
    )
  }, [
    filtered,
    currentPage
  ])

  async function uploadFile(event) {
    const file =
      event.target.files?.[0]

    if (!file) return

    setUploading(true)
    setError('')
    setMessage('Reading Excel file...')

    try {
      const buffer =
        await file.arrayBuffer()

      const workbook =
        XLSX.read(buffer, {
          type: 'array',
          cellDates: true
        })

      const worksheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ]

      const raw =
        XLSX.utils.sheet_to_json(
          worksheet,
          {
            defval: null,
            raw: false
          }
        )

      const mapped =
        raw
          .map(mapRow)
          .filter(
            item =>
              item.investor_name &&
              item.transaction_date &&
              item.amount !== null
          )

      if (!mapped.length) {
        throw new Error(
          'No valid transactions found in the Excel file.'
        )
      }

      setMessage(
        `Analysing ${mapped.length} transactions...`
      )

      const combinedRows = [
        ...rows,
        ...mapped
      ]

      const analysedUpload =
        mapped.map(item => {
          const classification =
            getClassification(
              item,
              combinedRows
            )

          return {
            ...item,
            classified_transaction_type:
              classification,
            classification_status:
              'Completed',
            classification_reason: null
          }
        })

      setMessage(
        `Uploading ${analysedUpload.length} transactions...`
      )

      for (
        let i = 0;
        i < analysedUpload.length;
        i += 500
      ) {
        const { error: uploadError } =
          await supabase
            .from('transactions')
            .insert(
              analysedUpload.slice(
                i,
                i + 500
              )
            )

        if (uploadError) {
          throw uploadError
        }
      }

      setMessage(
        'Excel analysed and uploaded successfully.'
      )

      await loadData()
      await loadRms()
    } catch (uploadError) {
      setError(
        uploadError.message ||
        'Unable to upload the file.'
      )

      setMessage('')
    }

    setUploading(false)
    event.target.value = ''
  }

  function exportExcel() {
    const output =
      filtered.map(item => ({
        Date: item.transaction_date,
        RM: item.rm_name,
        Investor: item.investor_name,
        Folio: item.folio_no,
        Scheme: item.scheme,
        Amount: item.amount,
        Source: item.source_type,
        Classification:
          item.system_classification
      }))

    const worksheet =
      XLSX.utils.json_to_sheet(output)

    const workbook =
      XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'Transaction Data'
    )

    XLSX.writeFile(
      workbook,
      'snowball-transaction-report.xlsx'
    )
  }

  function exportPDF() {
    const document =
      new jsPDF({
        orientation: 'landscape'
      })

    document.setFontSize(17)

    document.text(
      'Snowball Financial Services',
      14,
      14
    )

    document.setFontSize(11)

    document.text(
      'Transaction Data Report',
      14,
      21
    )

    autoTable(document, {
      startY: 28,

      head: [[
        'Date',
        'RM',
        'Investor',
        'Scheme',
        'Amount',
        'Source',
        'Classification'
      ]],

      body: filtered.map(item => [
        item.transaction_date,
        item.rm_name,
        item.investor_name,
        (item.scheme || '').slice(
          0,
          35
        ),
        money(item.amount),
        item.source_type,
        item.system_classification
      ])
    })

    document.save(
      'snowball-transaction-report.pdf'
    )
  }

  function classificationClass(value) {
    return String(value || '')
      .toLowerCase()
      .replace(/\s/g, '')
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
              Back to Login
            </button>

            <h1>
              Reset Password
            </h1>

            <p className="loginSub">
              Enter your email address to receive password reset instructions.
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
                value={forgotEmail}
                onChange={event =>
                  setForgotEmail(
                    event.target.value
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

              {forgotMessage && (
                <div className="successBox">
                  {forgotMessage}
                </div>
              )}

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

    return (
      <main className="loginPage">
        <section className="loginCard">

          <h1>
            Snowball Redemption Tracker
          </h1>

          <p className="loginSub">
            Sign in to monitor redemption activity
          </p>

          <form onSubmit={login}>

            <label>
              Email Address
            </label>

            <input
              type="email"
              placeholder="Enter your email"
              value={email}
              onChange={event =>
                setEmail(
                  event.target.value
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
                onChange={event =>
                  setPassword(
                    event.target.value
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
              Forgot Password?
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

  const navItems = [
    {
      id: 'dashboard',
      label: 'Dashboard',
      icon: LayoutDashboard
    },
    {
      id: 'transactions',
      label: 'Transaction Data',
      icon: Table2
    },
    {
      id: 'rms',
      label: 'Manage RMs',
      icon: Users
    }
  ]

  if (isAdmin) {
    navItems.push({
      id: 'settings',
      label: 'Settings',
      icon: SettingsIcon
    })
  }

  return (
    <div className="appShell">

      <aside className="sidebar">

        <div className="sidebarBrand">
          <img
            src={logo}
            alt="Snowball Financial Services"
          />
        </div>

        <nav className="sidebarNav">
          {navItems.map(item => {
            const Icon = item.icon

            return (
              <button
                key={item.id}
                className={
                  page === item.id
                    ? 'navActive'
                    : ''
                }
                onClick={() =>
                  setPage(item.id)
                }
              >
                <Icon size={18} />
                {item.label}
              </button>
            )
          })}
        </nav>

        <div className="sidebarBottom">
          <button
            onClick={logout}
          >
            <LogOut size={18} />
            Logout
          </button>
        </div>

      </aside>

      <main className="mainContent">

        <header className="topHeader">

          <div>
            <h1>
              {page === 'dashboard' &&
                'Dashboard'}

              {page === 'transactions' &&
                'Transaction Data'}

              {page === 'rms' &&
                'Manage Relationship Managers'}

              {page === 'settings' &&
                'Settings'}
            </h1>

            <p>
              {page === 'dashboard' &&
                'Analyse transactions and monitor redemption activity'}

              {page === 'transactions' &&
                'Upload and review detailed transaction data'}

              {page === 'rms' &&
                'Add, activate or mark Relationship Managers inactive'}

              {page === 'settings' &&
                'Manage application administrators'}
            </p>
          </div>

          <div className="headerActions">

            <button
              className="iconButton"
              onClick={() => {
                loadData()
                loadRms()
              }}
              title="Refresh"
            >
              <RefreshCw size={18} />
            </button>

            <div className="userBadge">
              {session.user.email
                ?.charAt(0)
                .toUpperCase()}
            </div>

          </div>

        </header>

        {error && (
          <div className="error globalMessage">
            <X
              size={16}
              onClick={() =>
                setError('')
              }
            />
            {error}
          </div>
        )}

        {message && (
          <div className="message globalMessage">
            {message}
          </div>
        )}

        {/* FILTER BAR ONLY FOR DASHBOARD AND TRANSACTION DATA */}

        {(
          page === 'dashboard' ||
          page === 'transactions'
        ) && (
          <section className="filterBar">

            <div className="periodButtons">
              {[
                'WTD',
                'MTD',
                'QTD',
                'YTD'
              ].map(item => (
                <button
                  key={item}
                  className={
                    period === item &&
                    !from &&
                    !to
                      ? 'active'
                      : ''
                  }
                  onClick={() => {
                    setPeriod(item)
                    setFrom('')
                    setTo('')
                  }}
                >
                  {item}
                </button>
              ))}
            </div>

            <select
              value={rm}
              onChange={event =>
                setRm(event.target.value)
              }
            >
              {activeRmNames.map(
                item => (
                  <option
                    key={item}
                    value={item}
                  >
                    {item === 'All'
                      ? 'All RMs'
                      : item}
                  </option>
                )
              )}
            </select>

            <input
              type="date"
              value={from}
              onChange={event =>
                setFrom(
                  event.target.value
                )
              }
            />

            <input
              type="date"
              value={to}
              onChange={event =>
                setTo(
                  event.target.value
                )
              }
            />

            <button
              className="refreshSmall"
              onClick={() => {
                setFrom('')
                setTo('')
                setPeriod('YTD')
                setRm('All')
              }}
              title="Reset filters"
            >
              <RefreshCw size={16} />
            </button>

          </section>
        )}

        {/* ================= DASHBOARD ================= */}

        {page === 'dashboard' && (
          <>

            <section className="summaryGrid">

              <article className="summaryCard redemptionCard">
                <span>
                  Redemption
                </span>
                <strong>
                  {money(
                    totals.Redemption
                  )}
                </strong>
              </article>

              <article className="summaryCard swpCard">
                <span>SWP</span>
                <strong>
                  {money(totals.SWP)}
                </strong>
              </article>

              <article className="summaryCard switchCard">
                <span>Switch</span>
                <strong>
                  {money(
                    totals.Switch
                  )}
                </strong>
              </article>

              <article className="summaryCard stpCard">
                <span>STP</span>
                <strong>
                  {money(totals.STP)}
                </strong>
              </article>

              <article className="summaryCard investorCard">
                <span>Investors</span>
                <strong>
                  {totals.Investors}
                </strong>
              </article>

              <article className="summaryCard transactionCard">
                <span>Transactions</span>
                <strong>
                  {totals.Transactions}
                </strong>
              </article>

            </section>

            <section className="dashboardGrid">

              <article className="chartCard classificationCard">

                <h2>
                  Amount by Classification
                </h2>

                <div className="donutArea">

                  <div
                    className="donut"
                    style={donutStyle}
                  >
                    <div className="donutHole">
                      <strong>
                        {totals.Transactions}
                      </strong>
                      <span>
                        Transactions
                      </span>
                    </div>
                  </div>

                  <div className="classificationLegend">

                    {classificationData.map(
                      item => (
                        <div
                          className={`legendRow ${classificationClass(
                            item.name
                          )}`}
                          key={item.name}
                        >
                          <span className="legendDot" />

                          <span>
                            {item.name}
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

              <article className="chartCard trendCard">

                <h2>
                  Monthly Trend (Amount in ₹)
                </h2>

                <div className="trendLegend">
                  <span className="redemption">
                    Redemption
                  </span>
                  <span className="swp">
                    SWP
                  </span>
                  <span className="switch">
                    Switch
                  </span>
                  <span className="stp">
                    STP
                  </span>
                </div>

                <div className="trendChart">

                  {monthlyTrend.map(
                    month => (
                      <div
                        className="monthGroup"
                        key={month.key}
                      >
                        <div className="bars">

                          {[
                            'Redemption',
                            'SWP',
                            'Switch',
                            'STP'
                          ].map(type => (
                            <div
                              key={type}
                              className={`miniBar ${classificationClass(
                                type
                              )}`}
                              style={{
                                height: `${
                                  Math.max(
                                    (
                                      month[type] /
                                      maxTrendAmount
                                    ) * 150,
                                    month[type]
                                      ? 3
                                      : 0
                                  )
                                }px`
                              }}
                              title={`${type}: ${money(
                                month[type]
                              )}`}
                            />
                          ))}

                        </div>

                        <span>
                          {month.label}
                        </span>
                      </div>
                    )
                  )}

                </div>

              </article>

            </section>

            <section className="dashboardGrid bottomCharts">

              <article className="chartCard rmChartCard">

                <h2>
                  Transactions by RM
                </h2>

                <div className="rmBars">

                  {rmChartData.map(
                    item => (
                      <div
                        className="rmBarRow"
                        key={item.name}
                      >
                        <span>
                          {item.name}
                        </span>

                        <div className="rmBarTrack">
                          <div
                            className="rmBarFill"
                            style={{
                              width: `${
                                (
                                  item.amount /
                                  maxRmAmount
                                ) * 100
                              }%`
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

                  {!rmChartData.length && (
                    <div className="emptyChart">
                      No data available
                    </div>
                  )}

                </div>

              </article>

              <article className="chartCard analysisCard">

                <h2>
                  Classification Analysis
                </h2>

                <div className="analysisRows">

                  {classificationData.map(
                    item => (
                      <div
                        className={`analysisRow ${classificationClass(
                          item.name
                        )}`}
                        key={item.name}
                      >
                        <div>
                          <span className="legendDot" />
                          {item.name}
                        </div>

                        <div className="analysisTrack">
                          <div
                            className="analysisFill"
                            style={{
                              width: `${
                                (
                                  item.amount /
                                  Math.max(
                                    ...classificationData.map(
                                      value =>
                                        value.amount
                                    ),
                                    1
                                  )
                                ) * 100
                              }%`
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

                </div>

              </article>

            </section>

          </>
        )}

        {/* ================= TRANSACTION DATA ================= */}

        {page === 'transactions' && (
          <section className="transactionPage">

            <div className="transactionActions">

              <button className="uploadButton">
                <Upload size={17} />

                <label>
                  {uploading
                    ? 'Uploading...'
                    : 'Upload Excel'}

                  <input
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    onChange={uploadFile}
                    disabled={uploading}
                  />
                </label>
              </button>

              <button
                onClick={exportExcel}
              >
                <Download size={17} />
                Excel
              </button>

              <button
                onClick={exportPDF}
              >
                <FileText size={17} />
                PDF
              </button>

            </div>

            <article className="tableCard">

              <div className="sectionHeading">
                <div>
                  <h2>
                    Transaction Details
                  </h2>

                  <p>
                    {filtered.length} transactions found
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
                      item => (
                        <tr key={item.id}>
                          <td>
                            {item.transaction_date}
                          </td>

                          <td>
                            {item.rm_name}
                          </td>

                          <td>
                            {item.investor_name}
                          </td>

                          <td>
                            {item.folio_no}
                          </td>

                          <td>
                            {item.scheme}
                          </td>

                          <td>
                            {money(item.amount)}
                          </td>

                          <td>
                            <span className="sourceBadge">
                              {item.source_type}
                            </span>
                          </td>

                          <td>
                            <span
                              className={`classification ${classificationClass(
                                item.system_classification
                              )}`}
                            >
                              {
                                item.system_classification
                              }
                            </span>
                          </td>
                        </tr>
                      )
                    )}

                    {!paginatedRows.length && (
                      <tr>
                        <td
                          colSpan="8"
                          className="emptyTable"
                        >
                          No transactions found.
                        </td>
                      </tr>
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

            </article>

          </section>
        )}

        {/* ================= MANAGE RMS ================= */}

        {page === 'rms' && (
          <section className="rmPage">

            <article className="rmAddCard">

              <h2>
                Add New RM
              </h2>

              <p>
                Add a new Relationship Manager to the organisation.
              </p>

              <form onSubmit={addRm}>

                <input
                  placeholder="Enter RM name"
                  value={newRm}
                  onChange={event =>
                    setNewRm(
                      event.target.value
                    )
                  }
                />

                <button className="primaryButton">
                  <UserPlus size={17} />
                  Add RM
                </button>

              </form>

            </article>

            <article className="rmListCard">

              <div className="rmTabs">

                <button
                  className={
                    rmView === 'current'
                      ? 'active'
                      : ''
                  }
                  onClick={() =>
                    setRmView('current')
                  }
                >
                  Current RMs (
                  {
                    rms.filter(
                      item =>
                        item.is_active
                    ).length
                  }
                  )
                </button>

                <button
                  className={
                    rmView === 'inactive'
                      ? 'active'
                      : ''
                  }
                  onClick={() =>
                    setRmView('inactive')
                  }
                >
                  Inactive RMs (
                  {
                    rms.filter(
                      item =>
                        !item.is_active
                    ).length
                  }
                  )
                </button>

                <button
                  className={
                    rmView === 'all'
                      ? 'active'
                      : ''
                  }
                  onClick={() =>
                    setRmView('all')
                  }
                >
                  All RMs
                </button>

              </div>

              <div className="rmList">

                {rms
                  .filter(item => {
                    if (
                      rmView === 'current'
                    ) {
                      return item.is_active
                    }

                    if (
                      rmView === 'inactive'
                    ) {
                      return !item.is_active
                    }

                    return true
                  })
                  .map(item => (
                    <div
                      className="rmRow"
                      key={item.id}
                    >

                      <div className="rmAvatar">
                        {item.name
                          ?.charAt(0)
                          .toUpperCase()}
                      </div>

                      <div className="rmName">

                        <strong>
                          {item.name}
                        </strong>

                        <span
                          className={
                            item.is_active
                              ? 'statusActive'
                              : 'statusInactive'
                          }
                        >
                          {item.is_active
                            ? 'Active'
                            : 'Inactive'}
                        </span>

                      </div>

                      <button
                        className={
                          item.is_active
                            ? 'deactivateButton'
                            : 'activateButton'
                        }
                        onClick={() =>
                          toggleRmStatus(
                            item
                          )
                        }
                      >
                        {item.is_active
                          ? (
                            <>
                              <UserMinus size={15} />
                              Mark Inactive
                            </>
                          )
                          : (
                            <>
                              <UserCheck size={15} />
                              Activate
                            </>
                          )}
                      </button>

                    </div>
                  ))}

                {!rms.length && (
                  <div className="emptyChart">
                    No RMs added yet.
                  </div>
                )}

              </div>

            </article>

          </section>
        )}

        {/* ================= SETTINGS ================= */}

        {page === 'settings' &&
          isAdmin && (
            <section className="settingsPage">

              <article className="settingsCard">

                <div className="settingsTitle">
                  <Shield size={21} />

                  <div>
                    <h2>
                      Admin Login Access
                    </h2>

                    <p>
                      Add email addresses that should have administrator access.
                    </p>
                  </div>
                </div>

                <form
                  className="adminForm"
                  onSubmit={addAdmin}
                >

                  <input
                    type="email"
                    placeholder="Enter admin email address"
                    value={newAdminEmail}
                    onChange={event =>
                      setNewAdminEmail(
                        event.target.value
                      )
                    }
                    required
                  />

                  <button
                    className="primaryButton"
                  >
                    <Plus size={17} />
                    Add Admin
                  </button>

                </form>

              </article>

              <article className="adminListCard">

                <h2>
                  Admin Users
                </h2>

                {adminEmails.map(
                  item => (
                    <div
                      className="adminRow"
                      key={item.id}
                    >

                      <div className="adminEmail">
                        <Mail size={17} />

                        <div>
                          <strong>
                            {item.email}
                          </strong>

                          <span
                            className={
                              item.is_active
                                ? 'statusActive'
                                : 'statusInactive'
                            }
                          >
                            {item.is_active
                              ? 'Active'
                              : 'Inactive'}
                          </span>
                        </div>
                      </div>

                      <button
                        className={
                          item.is_active
                            ? 'deactivateButton'
                            : 'activateButton'
                        }
                        onClick={() =>
                          toggleAdminStatus(
                            item
                          )
                        }
                      >
                        {item.is_active
                          ? 'Make Inactive'
                          : 'Activate'}
                      </button>

                    </div>
                  )
                )}

                {!adminEmails.length && (
                  <div className="emptyChart">
                    No admin emails found.
                  </div>
                )}

              </article>

            </section>
          )}

      </main>
    </div>
  )
}

createRoot(
  document.getElementById('root')
).render(<App />)
