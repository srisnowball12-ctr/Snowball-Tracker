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
  Upload,
  Download,
  FileText,
  LogOut,
  RefreshCw,
  Plus,
  Eye,
  EyeOff,
  LockKeyhole,
  ArrowLeft,
  X
} from 'lucide-react'
import './styles.css'

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
)

/* =========================================================
   HELPERS
========================================================= */

const money = value =>
  new Intl.NumberFormat('en-IN', {
    style: 'currency',
    currency: 'INR',
    maximumFractionDigits: 0
  }).format(Number(value || 0))

const shortMoney = value => {
  const n = Number(value || 0)

  if (Math.abs(n) >= 10000000) {
    return `₹${(n / 10000000).toFixed(2)} Cr`
  }

  if (Math.abs(n) >= 100000) {
    return `₹${(n / 100000).toFixed(1)} L`
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

  if (value instanceof Date && !isNaN(value)) {
    return value.toISOString().slice(0, 10)
  }

  const date = new Date(value)

  return isNaN(date)
    ? null
    : date.toISOString().slice(0, 10)
}

function sourceLabel(value) {
  const text = String(value || '').trim().toLowerCase()

  if (text.includes('swp')) return 'SWP'
  if (text.includes('switch')) return 'Switch'
  if (text.includes('stp')) return 'STP'

  if (
    text === 'red' ||
    text.includes('redeem') ||
    text.includes('redemption')
  ) {
    return 'Redemption'
  }

  return value || 'Redemption'
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
      .find(
        value =>
          value !== undefined &&
          value !== null &&
          value !== ''
      )

  const amountRaw = get(
    'Amount(₹)',
    'Amount',
    'amount'
  )

  const amount =
    typeof amountRaw === 'number'
      ? amountRaw
      : Number(
          String(amountRaw || '').replace(
            /[₹,\s]/g,
            ''
          )
        )

  return {
    rm_name:
      get(
        'Partner/Employee',
        'RM',
        'rm_name'
      ) || null,

    group_name:
      get('Group', 'group_name') || null,

    investor_name:
      get(
        'Investor',
        'Investor Name',
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
      get(
        'Scheme',
        'Fund',
        'scheme'
      ) || null,

    amount:
      Number.isFinite(amount)
        ? amount
        : null,

    transaction_type: 'Imported',

    original_transaction_type:
      get(
        'Type',
        'Transaction Type',
        'original_transaction_type'
      ) || 'Redemption',

    classified_transaction_type: null,

    classification_status: 'Auto Classified'
  }
}

function daysBetween(a, b) {
  if (!a || !b) return 9999

  const first = new Date(a + 'T00:00:00')
  const second = new Date(b + 'T00:00:00')

  return Math.round(
    Math.abs(
      first.getTime() - second.getTime()
    ) /
      (1000 * 60 * 60 * 24)
  )
}

/* =========================================================
   AUTOMATIC TRANSACTION ANALYSIS

   The source Excel may say RED / Redemption.
   The system analyses recurring withdrawals for the same:

   Investor + Folio + Scheme

   and can classify them as SWP.
========================================================= */

function analyseTransactions(records) {
  const groups = {}

  records.forEach((row, index) => {
    const key = [
      norm(row.investor_name),
      norm(row.folio_no),
      norm(row.scheme)
    ].join('|')

    if (!groups[key]) {
      groups[key] = []
    }

    groups[key].push({
      ...row,
      __index: index
    })
  })

  Object.values(groups).forEach(group => {
    group.sort(
      (a, b) =>
        new Date(a.transaction_date) -
        new Date(b.transaction_date)
    )

    group.forEach((row, position) => {
      const raw = String(
        row.original_transaction_type || ''
      ).toLowerCase()

      let classification = 'Redemption'

      /* Explicit transaction types */

      if (raw.includes('swp')) {
        classification = 'SWP'
      } else if (raw.includes('switch')) {
        classification = 'Switch'
      } else if (raw.includes('stp')) {
        classification = 'STP'
      } else {
        /*
          Analyse recurring RED transactions.
        */

        const intervals = []

        if (position > 0) {
          intervals.push(
            daysBetween(
              row.transaction_date,
              group[position - 1].transaction_date
            )
          )
        }

        if (position < group.length - 1) {
          intervals.push(
            daysBetween(
              group[position + 1].transaction_date,
              row.transaction_date
            )
          )
        }

        const monthlyIntervals =
          intervals.filter(
            days =>
              days >= 20 &&
              days <= 40
          ).length

        const quarterlyIntervals =
          intervals.filter(
            days =>
              days >= 75 &&
              days <= 105
          ).length

        /*
          Strong recurring pattern.
        */

        if (
          group.length >= 3 &&
          (
            monthlyIntervals >= 1 ||
            quarterlyIntervals >= 1
          )
        ) {
          classification = 'SWP'
        }

        /*
          Two approximately monthly withdrawals
          with reasonably similar amounts.
        */

        else if (
          group.length === 2 &&
          monthlyIntervals >= 1
        ) {
          const other =
            position === 0
              ? group[1]
              : group[0]

          const currentAmount =
            Math.abs(
              Number(row.amount || 0)
            )

          const otherAmount =
            Math.abs(
              Number(other.amount || 0)
            )

          const average =
            (currentAmount + otherAmount) / 2

          const difference =
            Math.abs(
              currentAmount - otherAmount
            )

          if (
            average > 0 &&
            difference / average <= 0.5
          ) {
            classification = 'SWP'
          }
        }
      }

      row.classified_transaction_type =
        classification

      row.classification_status =
        'Auto Classified'
    })
  })

  const output =
    new Array(records.length)

  Object.values(groups).forEach(group => {
    group.forEach(row => {
      const { __index, ...cleanRow } = row

      output[__index] = cleanRow
    })
  })

  return output
}

/* =========================================================
   SIMPLE SVG BAR CHART
========================================================= */

function BarChart({
  data,
  title,
  valueFormatter = shortMoney
}) {
  const max =
    Math.max(
      ...data.map(x => Number(x.value || 0)),
      1
    )

  return (
    <div className="chartCard">
      <h3>{title}</h3>

      {!data.length ? (
        <div className="emptyChart">
          No data available
        </div>
      ) : (
        <div className="barChart">
          {data.map((item, index) => (
            <div
              className="barItem"
              key={`${item.label}-${index}`}
            >
              <div className="barValue">
                {valueFormatter(item.value)}
              </div>

              <div className="barTrack">
                <div
                  className="barFill"
                  style={{
                    height: `${
                      Math.max(
                        4,
                        (Number(item.value || 0) /
                          max) *
                          100
                      )
                    }%`
                  }}
                />
              </div>

              <div
                className="barLabel"
                title={item.label}
              >
                {item.label}
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

/* =========================================================
   APP
========================================================= */

function App() {
  const [session, setSession] =
    useState(null)

  const [email, setEmail] =
    useState('')

  const [password, setPassword] =
    useState('')

  const [showPassword, setShowPassword] =
    useState(false)

  const [forgotMode, setForgotMode] =
    useState(false)

  const [resetEmail, setResetEmail] =
    useState('')

  const [resetMessage, setResetMessage] =
    useState('')

  const [error, setError] =
    useState('')

  const [loading, setLoading] =
    useState(false)

  const [rows, setRows] =
    useState([])

  const [rms, setRms] =
    useState([])

  const [rm, setRm] =
    useState('All')

  const [from, setFrom] =
    useState('')

  const [to, setTo] =
    useState('')

  const [period, setPeriod] =
    useState('YTD')

  const [uploading, setUploading] =
    useState(false)

  const [message, setMessage] =
    useState('')

  const [activeTab, setActiveTab] =
    useState('dashboard')

  const [newRmName, setNewRmName] =
    useState('')

  const [rmMessage, setRmMessage] =
    useState('')

  const [transactionPage, setTransactionPage] =
    useState(1)

  const pageSize = 50

  /* Authentication */

  useEffect(() => {
    supabase.auth
      .getSession()
      .then(({ data }) => {
        setSession(data.session)
      })

    const {
      data: { subscription }
    } =
      supabase.auth.onAuthStateChange(
        (_event, currentSession) => {
          setSession(currentSession)
        }
      )

    return () =>
      subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (session) {
      loadData()
      loadRms()
    }
  }, [session])

  /* =====================================================
     LOGIN
  ===================================================== */

  async function login(e) {
    e.preventDefault()

    setLoading(true)
    setError('')

    const { error } =
      await supabase.auth.signInWithPassword({
        email,
        password
      })

    if (error) {
      setError(error.message)
    }

    setLoading(false)
  }

  /* =====================================================
     FORGOT PASSWORD
  ===================================================== */

  async function resetPassword(e) {
    e.preventDefault()

    setError('')
    setResetMessage('')
    setLoading(true)

    const redirectTo =
      `${window.location.origin}/`

    const { error } =
      await supabase.auth.resetPasswordForEmail(
        resetEmail,
        {
          redirectTo
        }
      )

    if (error) {
      setError(error.message)
    } else {
      setResetMessage(
        'If this email is registered, a password reset link has been sent.'
      )
    }

    setLoading(false)
  }

  /* =====================================================
     LOAD TRANSACTIONS
  ===================================================== */

  async function loadData() {
    setLoading(true)
    setError('')

    const { data, error } =
      await supabase
        .from('transactions')
        .select('*')
        .order(
          'transaction_date',
          { ascending: false }
        )

    if (error) {
      setError(error.message)
    } else {
      setRows(data || [])
    }

    setLoading(false)
  }

  /* =====================================================
     LOAD RMS

     Uses the new rms table.
     If that table is not yet created, it still derives
     RM names from transactions temporarily.
  ===================================================== */

  async function loadRms() {
    const { data, error } =
      await supabase
        .from('rms')
        .select('*')
        .order('rm_name')

    if (!error) {
      setRms(data || [])
      return
    }

    const { data: transactionData } =
      await supabase
        .from('transactions')
        .select('rm_name')

    const derived =
      Array.from(
        new Set(
          (transactionData || [])
            .map(x => x.rm_name)
            .filter(Boolean)
        )
      )
        .sort()
        .map(name => ({
          id: name,
          rm_name: name,
          is_active: true,
          derived: true
        }))

    setRms(derived)
  }

  const activeRmNames =
    useMemo(
      () => [
        'All',
        ...rms
          .filter(
            x => x.is_active !== false
          )
          .map(x => x.rm_name)
      ],
      [rms]
    )

  /* =====================================================
     FILTER DATA
  ===================================================== */

  const filtered = useMemo(() => {
    return rows.filter(row => {
      if (
        rm !== 'All' &&
        row.rm_name !== rm
      ) {
        return false
      }

      const date =
        row.transaction_date

      if (from && date < from) {
        return false
      }

      if (to && date > to) {
        return false
      }

      if (
        !from &&
        !to &&
        date
      ) {
        const now = new Date()

        const transactionDate =
          new Date(
            date + 'T00:00:00'
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
            0,
            0,
            0,
            0
          )

          if (
            transactionDate < start
          ) {
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
          const currentQuarter =
            Math.floor(
              now.getMonth() / 3
            )

          const transactionQuarter =
            Math.floor(
              transactionDate.getMonth() / 3
            )

          if (
            transactionDate.getFullYear() !==
              now.getFullYear() ||
            transactionQuarter !==
              currentQuarter
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
    rows,
    rm,
    from,
    to,
    period
  ])

  /* =====================================================
     DASHBOARD TOTALS
  ===================================================== */

  const totals = useMemo(() => {
    const sumType = type =>
      filtered
        .filter(
          x =>
            x.classified_transaction_type ===
            type
        )
        .reduce(
          (sum, x) =>
            sum +
            Number(x.amount || 0),
          0
        )

    return {
      Redemption:
        sumType('Redemption'),

      SWP:
        sumType('SWP'),

      Switch:
        sumType('Switch'),

      STP:
        sumType('STP'),

      Investors:
        new Set(
          filtered
            .map(
              x => x.investor_name
            )
            .filter(Boolean)
        ).size,

      Transactions:
        filtered.length
    }
  }, [filtered])

  /* =====================================================
     CHART DATA
  ===================================================== */

  const classificationData =
    useMemo(
      () => [
        {
          label: 'Redemption',
          value: totals.Redemption
        },
        {
          label: 'SWP',
          value: totals.SWP
        },
        {
          label: 'Switch',
          value: totals.Switch
        },
        {
          label: 'STP',
          value: totals.STP
        }
      ],
      [totals]
    )

  const monthlyData =
    useMemo(() => {
      const map = {}

      filtered.forEach(row => {
        if (!row.transaction_date) return

        const date =
          new Date(
            row.transaction_date +
              'T00:00:00'
          )

        const key =
          `${date.getFullYear()}-${String(
            date.getMonth() + 1
          ).padStart(2, '0')}`

        if (!map[key]) {
          map[key] = 0
        }

        map[key] +=
          Number(row.amount || 0)
      })

      return Object.entries(map)
        .sort(([a], [b]) =>
          a.localeCompare(b)
        )
        .slice(-12)
        .map(([key, value]) => {
          const [year, month] =
            key.split('-')

          const label =
            new Date(
              Number(year),
              Number(month) - 1,
              1
            ).toLocaleString(
              'en-IN',
              {
                month: 'short'
              }
            )

          return {
            label,
            value
          }
        })
    }, [filtered])

  const rmData =
    useMemo(() => {
      const map = {}

      filtered.forEach(row => {
        const name =
          row.rm_name || 'Unassigned'

        if (!map[name]) {
          map[name] = 0
        }

        map[name] +=
          Number(row.amount || 0)
      })

      return Object.entries(map)
        .sort(
          (a, b) => b[1] - a[1]
        )
        .slice(0, 8)
        .map(([label, value]) => ({
          label,
          value
        }))
    }, [filtered])

  /* =====================================================
     UPLOAD EXCEL
  ===================================================== */

  async function uploadFile(e) {
    const file =
      e.target.files?.[0]

    if (!file) return

    setUploading(true)
    setError('')
    setMessage('Reading Excel...')

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
            defval: null
          }
        )

      const mapped = raw
        .map(mapRow)
        .filter(
          row =>
            row.investor_name &&
            row.transaction_date &&
            row.amount !== null
        )

      if (!mapped.length) {
        throw new Error(
          'No valid transactions found. Please use the normal Snowball redemption Excel format.'
        )
      }

      setMessage(
        'Analysing SWP and Redemption patterns...'
      )

      /*
        Analyse existing data together with
        the newly uploaded transactions.
      */

      const combined = [
        ...rows,
        ...mapped
      ]

      const analysed =
        analyseTransactions(combined)

      const newRows =
        analysed.slice(rows.length)

      setMessage(
        `Uploading ${newRows.length} analysed transactions...`
      )

      for (
        let i = 0;
        i < newRows.length;
        i += 500
      ) {
        const batch =
          newRows.slice(
            i,
            i + 500
          )

        const { error } =
          await supabase
            .from('transactions')
            .insert(batch)

        if (error) {
          throw error
        }
      }

      setMessage(
        'Upload complete. Dashboard updated.'
      )

      await loadData()
      await loadRms()

      setActiveTab('dashboard')
    } catch (err) {
      setError(err.message)
      setMessage('')
    }

    setUploading(false)
    e.target.value = ''
  }

  /* =====================================================
     RE-ANALYSE ALL EXISTING DATA
  ===================================================== */

  async function reanalyseAll() {
    if (!rows.length) return

    setUploading(true)
    setError('')
    setMessage(
      'Re-analysing existing transactions...'
    )

    try {
      const analysed =
        analyseTransactions(rows)

      for (
        let i = 0;
        i < analysed.length;
        i += 100
      ) {
        const batch =
          analysed.slice(
            i,
            i + 100
          )

        await Promise.all(
          batch.map(row =>
            supabase
              .from('transactions')
              .update({
                classified_transaction_type:
                  row.classified_transaction_type,

                classification_status:
                  'Auto Classified'
              })
              .eq('id', row.id)
          )
        )
      }

      setMessage(
        'Re-analysis complete.'
      )

      await loadData()
    } catch (err) {
      setError(err.message)
      setMessage('')
    }

    setUploading(false)
  }

  /* =====================================================
     RM MANAGEMENT
  ===================================================== */

  async function addRm(e) {
    e.preventDefault()

    const name =
      newRmName.trim()

    if (!name) return

    setRmMessage('')

    const { error } =
      await supabase
        .from('rms')
        .insert({
          rm_name: name,
          is_active: true
        })

    if (error) {
      setRmMessage(
        error.message
      )
    } else {
      setNewRmName('')
      setRmMessage(
        `${name} added successfully.`
      )

      await loadRms()
    }
  }

  async function toggleRm(
    rmRecord
  ) {
    if (rmRecord.derived) {
      setRmMessage(
        'Please run the RM database setup first. Historical RMs are currently derived from transactions.'
      )
      return
    }

    const { error } =
      await supabase
        .from('rms')
        .update({
          is_active:
            !rmRecord.is_active
        })
        .eq(
          'id',
          rmRecord.id
        )

    if (error) {
      setRmMessage(
        error.message
      )
    } else {
      await loadRms()
    }
  }

  /* =====================================================
     EXPORT EXCEL
  ===================================================== */

  function exportExcel() {
    const output =
      filtered.map(row => ({
        Date:
          row.transaction_date,

        RM:
          row.rm_name,

        Investor:
          row.investor_name,

        Folio:
          row.folio_no,

        Scheme:
          row.scheme,

        Amount:
          row.amount,

        Source:
          sourceLabel(
            row.original_transaction_type
          ),

        Classification:
          row.classified_transaction_type
      }))

    const worksheet =
      XLSX.utils.json_to_sheet(
        output
      )

    const workbook =
      XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'Redemption Tracker'
    )

    XLSX.writeFile(
      workbook,
      'snowball-redemption-report.xlsx'
    )
  }

  /* =====================================================
     EXPORT PDF
  ===================================================== */

  function exportPDF() {
    const doc = new jsPDF({
      orientation: 'landscape'
    })

    doc.setFontSize(16)

    doc.text(
      'Snowball Financial Services - Redemption Tracker',
      14,
      14
    )

    doc.setFontSize(10)

    doc.text(
      `RM: ${rm} | Period: ${
        from || to
          ? `${from || ''} to ${to || ''}`
          : period
      }`,
      14,
      21
    )

    autoTable(doc, {
      startY: 27,

      head: [[
        'Date',
        'RM',
        'Investor',
        'Scheme',
        'Amount',
        'Source',
        'Classification'
      ]],

      body:
        filtered.map(row => [
          row.transaction_date,
          row.rm_name,
          row.investor_name,
          (row.scheme || '')
            .slice(0, 35),
          money(row.amount),
          sourceLabel(
            row.original_transaction_type
          ),
          row.classified_transaction_type
        ])
    })

    doc.save(
      'snowball-redemption-report.pdf'
    )
  }

  /* =====================================================
     LOGIN SCREEN
  ===================================================== */

  if (!session) {
    return (
      <main className="loginPage">
        <section className="loginCard">

          <div className="loginBrand">
            <div className="logoCircle">
              S
            </div>

            <div>
              <h1>
                Snowball Financial Services
              </h1>

              <p>
                Redemption Tracker
              </p>
            </div>
          </div>

          {!forgotMode ? (
            <>
              <h2>
                Welcome back
              </h2>

              <p className="loginSub">
                Sign in to access your
                dashboard
              </p>

              <form onSubmit={login}>
                <label>
                  Email
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
                    placeholder="Enter password"
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
                      : <Eye size={18} />
                    }
                  </button>
                </div>

                <button
                  className="primaryButton fullButton"
                  disabled={loading}
                >
                  {loading
                    ? 'Signing in...'
                    : 'Login'
                  }
                </button>
              </form>

              <button
                className="forgotLink"
                onClick={() => {
                  setForgotMode(true)
                  setError('')
                  setResetMessage('')
                }}
              >
                Forgot / Reset Password?
              </button>
            </>
          ) : (
            <>
              <button
                className="backLink"
                onClick={() => {
                  setForgotMode(false)
                  setError('')
                  setResetMessage('')
                }}
              >
                <ArrowLeft size={17} />
                Back to Login
              </button>

              <h2>
                Reset Password
              </h2>

              <p className="loginSub">
                Enter your registered email
                address.
              </p>

              <form
                onSubmit={resetPassword}
              >
                <label>
                  Email
                </label>

                <input
                  type="email"
                  placeholder="Enter your email"
                  value={resetEmail}
                  onChange={e =>
                    setResetEmail(
                      e.target.value
                    )
                  }
                  required
                />

                <button
                  className="primaryButton fullButton"
                  disabled={loading}
                >
                  <LockKeyhole size={17} />

                  {loading
                    ? 'Sending...'
                    : 'Send Reset Link'}
                </button>
              </form>

              {resetMessage && (
                <div className="successBox">
                  {resetMessage}
                </div>
              )}
            </>
          )}

          {error && (
            <div className="error">
              {error}
            </div>
          )}

        </section>
      </main>
    )
  }

  /* =====================================================
     MAIN APPLICATION
  ===================================================== */

  const totalPages =
    Math.max(
      1,
      Math.ceil(
        filtered.length /
          pageSize
      )
    )

  const transactionRows =
    filtered.slice(
      (transactionPage - 1) *
        pageSize,
      transactionPage *
        pageSize
    )

  return (
    <div className="appShell">

      {/* SIDEBAR */}

      <aside className="sidebar">

        <div className="sidebarBrand">

          <div className="logoCircle">
            S
          </div>

          <div>
            <h2>
              Snowball
            </h2>

            <span>
              Financial Services
            </span>
          </div>

        </div>

        <nav className="sidebarNav">

          <button
            className={
              activeTab === 'dashboard'
                ? 'navActive'
                : ''
            }
            onClick={() =>
              setActiveTab(
                'dashboard'
              )
            }
          >
            <LayoutDashboard size={19} />
            Dashboard
          </button>

          <button
            className={
              activeTab === 'transactions'
                ? 'navActive'
                : ''
            }
            onClick={() => {
              setActiveTab(
                'transactions'
              )
              setTransactionPage(1)
            }}
          >
            <Table2 size={19} />
            Transaction Data
          </button>

          <button
            className={
              activeTab === 'rms'
                ? 'navActive'
                : ''
            }
            onClick={() =>
              setActiveTab('rms')
            }
          >
            <Users size={19} />
            RM Management
          </button>

        </nav>

        <div className="sidebarBottom">
          <button
            onClick={() =>
              supabase.auth.signOut()
            }
          >
            <LogOut size={18} />
            Logout
          </button>
        </div>

      </aside>

      {/* MAIN CONTENT */}

      <main className="mainContent">

        <header className="topHeader">

          <div>
            <h1>
              {
                activeTab === 'dashboard'
                  ? 'Dashboard'
                  : activeTab === 'transactions'
                    ? 'Transaction Data'
                    : 'RM Management'
              }
            </h1>

            <p>
              Snowball Redemption Tracker
            </p>
          </div>

          <div className="headerActions">

            <button
              className="iconButton"
              onClick={loadData}
              title="Refresh data"
            >
              <RefreshCw size={18} />
            </button>

            <div className="userBadge">
              {session.user.email
                ?.charAt(0)
                ?.toUpperCase()}
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

        {/* =================================================
            FILTER BAR
        ================================================= */}

        {activeTab !== 'rms' && (
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
                      ? 'activePeriod'
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
              onChange={e =>
                setRm(
                  e.target.value
                )
              }
            >
              {activeRmNames.map(
                name => (
                  <option
                    key={name}
                  >
                    {name}
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

            <button className="uploadButton">

              <Upload size={17} />

              <label>
                {uploading
                  ? 'Processing...'
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
              className="secondaryButton"
              onClick={exportExcel}
            >
              <Download size={17} />
              Excel
            </button>

            <button
              className="secondaryButton"
              onClick={exportPDF}
            >
              <FileText size={17} />
              PDF
            </button>

          </section>
        )}

        {/* =================================================
            DASHBOARD
        ================================================= */}

        {activeTab === 'dashboard' && (
          <>

            <section className="kpiGrid">

              <article className="kpiCard">
                <span>
                  Total Redemption
                </span>

                <strong>
                  {shortMoney(
                    totals.Redemption
                  )}
                </strong>
              </article>

              <article className="kpiCard">
                <span>
                  SWP
                </span>

                <strong>
                  {shortMoney(
                    totals.SWP
                  )}
                </strong>
              </article>

              <article className="kpiCard">
                <span>
                  Switch
                </span>

                <strong>
                  {shortMoney(
                    totals.Switch
                  )}
                </strong>
              </article>

              <article className="kpiCard">
                <span>
                  Investors
                </span>

                <strong>
                  {totals.Investors}
                </strong>
              </article>

              <article className="kpiCard">
                <span>
                  Transactions
                </span>

                <strong>
                  {totals.Transactions}
                </strong>
              </article>

            </section>

            <section className="chartsGrid">

              <BarChart
                title="Transaction Classification"
                data={classificationData}
              />

              <BarChart
                title="Monthly Redemption Trend"
                data={monthlyData}
              />

            </section>

            <section className="singleChart">
              <BarChart
                title="RM-wise Redemption Activity"
                data={rmData}
              />
            </section>

            <section className="recentCard">

              <div className="sectionHeading">

                <div>
                  <h2>
                    Recent Transactions
                  </h2>

                  <p>
                    Latest analysed
                    transactions
                  </p>
                </div>

                <button
                  className="viewAllButton"
                  onClick={() => {
                    setActiveTab(
                      'transactions'
                    )
                    setTransactionPage(1)
                  }}
                >
                  View All
                </button>

              </div>

              <div className="miniTableWrap">

                <table>

                  <thead>
                    <tr>
                      <th>Date</th>
                      <th>Investor</th>
                      <th>RM</th>
                      <th>Amount</th>
                      <th>Classification</th>
                    </tr>
                  </thead>

                  <tbody>

                    {filtered
                      .slice(0, 8)
                      .map(row => (
                        <tr
                          key={row.id}
                        >
                          <td>
                            {
                              row.transaction_date
                            }
                          </td>

                          <td>
                            {
                              row.investor_name
                            }
                          </td>

                          <td>
                            {
                              row.rm_name
                            }
                          </td>

                          <td>
                            {
                              money(
                                row.amount
                              )
                            }
                          </td>

                          <td>
                            <span
                              className={`classification ${
                                String(
                                  row.classified_transaction_type ||
                                    ''
                                )
                                  .toLowerCase()
                                  .replace(
                                    /\s/g,
                                    ''
                                  )
                              }`}
                            >
                              {
                                row.classified_transaction_type ||
                                  'Pending'
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

        {/* =================================================
            TRANSACTION DATA
        ================================================= */}

        {activeTab === 'transactions' && (
          <section className="dataCard">

            <div className="sectionHeading">

              <div>
                <h2>
                  All Transactions
                </h2>

                <p>
                  {filtered.length} records
                </p>
              </div>

              <button
                className="secondaryButton"
                onClick={reanalyseAll}
                disabled={uploading}
              >
                <RefreshCw size={17} />
                Re-analyse Data
              </button>

            </div>

            <div className="tableScroll">

              <table className="dataTable">

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

                  {transactionRows.map(
                    row => (
                      <tr
                        key={row.id}
                      >
                        <td>
                          {
                            row.transaction_date
                          }
                        </td>

                        <td>
                          {
                            row.rm_name
                          }
                        </td>

                        <td>
                          {
                            row.investor_name
                          }
                        </td>

                        <td>
                          {
                            row.folio_no
                          }
                        </td>

                        <td>
                          {
                            row.scheme
                          }
                        </td>

                        <td>
                          {
                            money(
                              row.amount
                            )
                          }
                        </td>

                        <td>
                          <span className="sourceBadge">
                            {sourceLabel(
                              row.original_transaction_type
                            )}
                          </span>
                        </td>

                        <td>
                          <span
                            className={`classification ${
                              String(
                                row.classified_transaction_type ||
                                  ''
                              )
                                .toLowerCase()
                                .replace(
                                  /\s/g,
                                  ''
                                )
                            }`}
                          >
                            {
                              row.classified_transaction_type ||
                                'Pending'
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
                  transactionPage === 1
                }
                onClick={() =>
                  setTransactionPage(
                    page =>
                      Math.max(
                        1,
                        page - 1
                      )
                  )
                }
              >
                Previous
              </button>

              <span>
                Page {transactionPage}
                {' '}of{' '}
                {totalPages}
              </span>

              <button
                disabled={
                  transactionPage ===
                  totalPages
                }
                onClick={() =>
                  setTransactionPage(
                    page =>
                      Math.min(
                        totalPages,
                        page + 1
                      )
                  )
                }
              >
                Next
              </button>

            </div>

          </section>
        )}

        {/* =================================================
            RM MANAGEMENT
        ================================================= */}

        {activeTab === 'rms' && (
          <section className="rmPage">

            <div className="rmAddCard">

              <h2>
                Add Relationship Manager
              </h2>

              <p>
                Add new RMs as your
                organisation grows.
              </p>

              <form onSubmit={addRm}>

                <input
                  placeholder="Enter RM name"
                  value={newRmName}
                  onChange={e =>
                    setNewRmName(
                      e.target.value
                    )
                  }
                />

                <button
                  className="primaryButton"
                >
                  <Plus size={17} />
                  Add RM
                </button>

              </form>

              {rmMessage && (
                <div className="message">
                  {rmMessage}
                </div>
              )}

            </div>

            <section className="rmListCard">

              <div className="sectionHeading">
                <div>
                  <h2>
                    Relationship Managers
                  </h2>

                  <p>
                    Activate or deactivate
                    RMs without losing
                    historical data.
                  </p>
                </div>
              </div>

              <div className="rmList">

                {rms.map(
                  rmRecord => (
                    <div
                      className="rmRow"
                      key={rmRecord.id}
                    >

                      <div className="rmAvatar">
                        {rmRecord.rm_name
                          ?.charAt(0)
                          ?.toUpperCase()}
                      </div>

                      <div className="rmName">
                        <strong>
                          {
                            rmRecord.rm_name
                          }
                        </strong>

                        <span
                          className={
                            rmRecord.is_active !==
                            false
                              ? 'statusActive'
                              : 'statusInactive'
                          }
                        >
                          {rmRecord.is_active !==
                          false
                            ? 'Active'
                            : 'Inactive'}
                        </span>
                      </div>

                      <button
                        className={
                          rmRecord.is_active !==
                          false
                            ? 'deactivateButton'
                            : 'activateButton'
                        }
                        onClick={() =>
                          toggleRm(
                            rmRecord
                          )
                        }
                      >
                        {rmRecord.is_active !==
                        false
                          ? 'Deactivate'
                          : 'Reactivate'}
                      </button>

                    </div>
                  )
                )}

              </div>

            </section>

          </section>
        )}

      </main>

    </div>
  )
}

createRoot(
  document.getElementById('root')
).render(<App />)
