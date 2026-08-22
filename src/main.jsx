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
  Trash2,
  KeyRound,
  ArrowLeft,
  X
} from 'lucide-react'
import './styles.css'

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
)

/* -------------------------------------------------------
   HELPERS
------------------------------------------------------- */

const money = n =>
  new Intl.NumberFormat('en-IN', {
    style: 'currency',
    currency: 'INR',
    maximumFractionDigits: 0
  }).format(Number(n || 0))

const shortMoney = n => {
  const value = Number(n || 0)

  if (value >= 10000000) {
    return `₹${(value / 10000000).toFixed(1)}Cr`
  }

  if (value >= 100000) {
    return `₹${(value / 100000).toFixed(1)}L`
  }

  if (value >= 1000) {
    return `₹${(value / 1000).toFixed(1)}K`
  }

  return money(value)
}

const iso = value => {
  if (!value) return null

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10)
  }

  const date = new Date(value)

  return isNaN(date)
    ? null
    : date.toISOString().slice(0, 10)
}

const norm = value =>
  String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')

const normaliseType = value => {
  const v = String(value || '').toLowerCase().trim()

  if (v.includes('swp')) return 'SWP'
  if (v.includes('stp')) return 'STP'
  if (v.includes('switch')) return 'Switch'

  if (
    v === 'red' ||
    v.includes('redemption') ||
    v.includes('redeem') ||
    v.includes('red ')
  ) {
    return 'Redemption'
  }

  return ''
}

const getClassification = row => {
  return (
    normaliseType(row.classified_transaction_type) ||
    normaliseType(row.transaction_type) ||
    'Redemption'
  )
}

const getSource = row => {
  return (
    normaliseType(row.original_transaction_type) ||
    normaliseType(row.transaction_type) ||
    normaliseType(row.classified_transaction_type) ||
    '-'
  )
}

/* -------------------------------------------------------
   EXCEL MAPPING
------------------------------------------------------- */

function mapRow(row) {
  const lookup = Object.fromEntries(
    Object.entries(row).map(([key, value]) => [norm(key), value])
  )

  const get = (...keys) =>
    keys
      .map(key => lookup[norm(key)])
      .find(value => value !== undefined)

  const amountRaw = get('Amount(₹)', 'Amount', 'amount')

  const amount =
    typeof amountRaw === 'number'
      ? amountRaw
      : Number(String(amountRaw || '').replace(/[₹,\s]/g, ''))

  return {
    rm_name:
      get('Partner/Employee', 'RM', 'rm_name', 'Partner') || null,

    group_name:
      get('Group', 'group_name') || null,

    investor_name:
      get('Investor', 'investor_name', 'Investor Name') || null,

    transaction_date:
      iso(get('Date', 'transaction_date', 'Transaction Date')),

    folio_no:
      String(
        get(
          'Folio No/Demat A/C',
          'Folio',
          'folio_no',
          'Folio Number'
        ) || ''
      ) || null,

    scheme:
      get('Scheme', 'Fund', 'scheme', 'Scheme Name') || null,

    amount:
      Number.isFinite(amount)
        ? amount
        : null,

    transaction_type:
      'Imported',

    original_transaction_type:
      get('Type', 'Transaction Type', 'original_transaction_type') || null,

    classification_status:
      'Needs Review',

    classification_reason:
      null
  }
}

/* -------------------------------------------------------
   SMALL CHART COMPONENTS
------------------------------------------------------- */

function DonutChart({ data }) {
  const total = data.reduce((sum, item) => sum + item.value, 0)

  if (!total) {
    return (
      <div className="emptyChart">
        No transaction data
      </div>
    )
  }

  let cumulative = 0

  const segments = data.map(item => {
    const start = cumulative
    const percent = (item.value / total) * 100
    cumulative += percent

    return `${item.color} ${start}% ${cumulative}%`
  })

  return (
    <div className="donutSection">

      <div
        className="donut"
        style={{
          background: `conic-gradient(${segments.join(', ')})`
        }}
      >
        <div className="donutCenter">
          <strong>{data.length}</strong>
          <span>Types</span>
        </div>
      </div>

      <div className="legend">
        {data.map(item => (
          <div className="legendItem" key={item.label}>
            <span
              className="legendDot"
              style={{ background: item.color }}
            />
            <span>{item.label}</span>
            <b>{shortMoney(item.value)}</b>
          </div>
        ))}
      </div>

    </div>
  )
}

function MonthlyTrend({ rows }) {
  const months = []
  const now = new Date()

  for (let i = 5; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1)

    months.push({
      month: d.toLocaleString('en-IN', { month: 'short' }),
      monthNumber: d.getMonth(),
      year: d.getFullYear()
    })
  }

  const series = ['Redemption', 'SWP', 'Switch', 'STP'].map(type => ({
    type,
    values: months.map(month =>
      rows
        .filter(row => {
          if (!row.transaction_date) return false

          const d = new Date(row.transaction_date + 'T00:00:00')

          return (
            d.getMonth() === month.monthNumber &&
            d.getFullYear() === month.year &&
            getClassification(row) === type
          )
        })
        .reduce((sum, row) => sum + Number(row.amount || 0), 0)
    )
  }))

  const maxValue = Math.max(
    ...series.flatMap(s => s.values),
    1
  )

  const colors = {
    Redemption: '#c9404c',
    SWP: '#2f9c75',
    Switch: '#3f7edb',
    STP: '#8d65d8'
  }

  const width = 520
  const height = 220
  const left = 25
  const right = 15
  const top = 20
  const bottom = 35

  const plotWidth = width - left - right
  const plotHeight = height - top - bottom

  const pointsFor = values =>
    values
      .map((value, index) => {
        const x =
          left +
          (index / Math.max(values.length - 1, 1)) *
            plotWidth

        const y =
          top +
          plotHeight -
          (value / maxValue) * plotHeight

        return `${x},${y}`
      })
      .join(' ')

  return (
    <div className="trendWrap">

      <svg
        viewBox={`0 0 ${width} ${height}`}
        className="trendChart"
      >
        {[0, 0.25, 0.5, 0.75, 1].map((value, index) => {
          const y = top + value * plotHeight

          return (
            <line
              key={index}
              x1={left}
              x2={width - right}
              y1={y}
              y2={y}
              className="gridLine"
            />
          )
        })}

        {series.map(seriesItem => (
          <polyline
            key={seriesItem.type}
            fill="none"
            stroke={colors[seriesItem.type]}
            strokeWidth="3"
            strokeLinecap="round"
            strokeLinejoin="round"
            points={pointsFor(seriesItem.values)}
          />
        ))}

        {series.map(seriesItem =>
          seriesItem.values.map((value, index) => {
            const x =
              left +
              (index /
                Math.max(seriesItem.values.length - 1, 1)) *
                plotWidth

            const y =
              top +
              plotHeight -
              (value / maxValue) * plotHeight

            return (
              <circle
                key={`${seriesItem.type}-${index}`}
                cx={x}
                cy={y}
                r="3.5"
                fill={colors[seriesItem.type]}
              />
            )
          })
        )}

        {months.map((month, index) => {
          const x =
            left +
            (index / Math.max(months.length - 1, 1)) *
              plotWidth

          return (
            <text
              key={month.month}
              x={x}
              y={height - 10}
              textAnchor="middle"
              className="axisText"
            >
              {month.month}
            </text>
          )
        })}
      </svg>

      <div className="chartLegend">
        <span><i className="dot red" />Redemption</span>
        <span><i className="dot green" />SWP</span>
        <span><i className="dot blue" />Switch</span>
        <span><i className="dot purple" />STP</span>
      </div>

    </div>
  )
}

function RMBarChart({ rows }) {
  const data = useMemo(() => {
    const grouped = {}

    rows.forEach(row => {
      const name = row.rm_name || 'Unassigned'

      grouped[name] =
        (grouped[name] || 0) +
        Number(row.amount || 0)
    })

    return Object.entries(grouped)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 6)
  }, [rows])

  const maxValue = Math.max(
    ...data.map(item => item.value),
    1
  )

  if (!data.length) {
    return (
      <div className="emptyChart">
        No RM data available
      </div>
    )
  }

  return (
    <div className="rmBars">
      {data.map(item => (
        <div className="rmBarRow" key={item.name}>

          <span className="rmBarName">
            {item.name}
          </span>

          <div className="rmBarTrack">
            <div
              className="rmBarFill"
              style={{
                width: `${Math.max(
                  (item.value / maxValue) * 100,
                  2
                )}%`
              }}
            />
          </div>

          <b>{shortMoney(item.value)}</b>

        </div>
      ))}
    </div>
  )
}

/* -------------------------------------------------------
   APP
------------------------------------------------------- */

function App() {
  const [session, setSession] = useState(null)

  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [message, setMessage] = useState('')
  const [loading, setLoading] = useState(false)

  const [resetMode, setResetMode] = useState(false)
  const [resetSent, setResetSent] = useState(false)

  const [activeTab, setActiveTab] =
    useState('dashboard')

  const [rows, setRows] = useState([])
  const [rms, setRms] = useState([])

  const [rm, setRm] = useState('All')
  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')
  const [period, setPeriod] = useState('YTD')

  const [uploading, setUploading] = useState(false)

  const [newRM, setNewRM] = useState('')
  const [rmMessage, setRmMessage] = useState('')

  /* ---------------- AUTH ---------------- */

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
      }
    )

    return () => subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (session) {
      loadData()
      loadRMs()
    }
  }, [session])

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

  async function sendPasswordReset(e) {
    e.preventDefault()

    setLoading(true)
    setError('')
    setResetSent(false)

    const { error } =
      await supabase.auth.resetPasswordForEmail(
        email,
        {
          redirectTo: window.location.origin
        }
      )

    if (error) {
      setError(error.message)
    } else {
      setResetSent(true)
    }

    setLoading(false)
  }

  async function logout() {
    await supabase.auth.signOut()
    setActiveTab('dashboard')
  }

  /* ---------------- DATA ---------------- */

  async function loadData() {
    setLoading(true)

    const { data, error } = await supabase
      .from('transactions')
      .select('*')
      .order('transaction_date', {
        ascending: false
      })

    if (error) {
      setError(error.message)
    } else {
      setRows(data || [])
    }

    setLoading(false)
  }

  async function loadRMs() {
    const { data, error } = await supabase
      .from('rms')
      .select('*')
      .order('name')

    if (!error && data) {
      setRms(data)
    } else {
      const names = Array.from(
        new Set(
          rows
            .map(row => row.rm_name)
            .filter(Boolean)
        )
      ).sort()

      setRms(
        names.map(name => ({
          id: name,
          name,
          fallback: true
        }))
      )
    }
  }

  /* ---------------- FILTERS ---------------- */

  const filtered = useMemo(() => {
    return rows.filter(row => {
      if (
        rm !== 'All' &&
        row.rm_name !== rm
      ) {
        return false
      }

      const date = row.transaction_date

      if (!date) return false

      if (from && date < from) {
        return false
      }

      if (to && date > to) {
        return false
      }

      if (!from && !to) {
        const now = new Date()
        const dt = new Date(
          date + 'T00:00:00'
        )

        if (period === 'WTD') {
          const day =
            (now.getDay() + 6) % 7

          const start = new Date(now)

          start.setDate(
            now.getDate() - day
          )

          start.setHours(0, 0, 0, 0)

          if (dt < start) return false
        }

        if (
          period === 'MTD' &&
          (
            dt.getMonth() !== now.getMonth() ||
            dt.getFullYear() !== now.getFullYear()
          )
        ) {
          return false
        }

        if (period === 'QTD') {
          const quarter =
            Math.floor(now.getMonth() / 3)

          if (
            dt.getFullYear() !==
              now.getFullYear() ||
            Math.floor(
              dt.getMonth() / 3
            ) !== quarter
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
    rows,
    rm,
    from,
    to,
    period
  ])

  /* ---------------- TOTALS ---------------- */

  const totals = useMemo(() => {
    const calculate = type =>
      filtered
        .filter(
          row =>
            getClassification(row) === type
        )
        .reduce(
          (sum, row) =>
            sum +
            Number(row.amount || 0),
          0
        )

    return {
      Redemption:
        calculate('Redemption'),

      SWP:
        calculate('SWP'),

      Switch:
        calculate('Switch'),

      STP:
        calculate('STP'),

      Investors:
        new Set(
          filtered
            .map(
              row => row.investor_name
            )
            .filter(Boolean)
        ).size,

      Transactions:
        filtered.length
    }
  }, [filtered])

  const classificationData = [
    {
      label: 'Redemption',
      value: totals.Redemption,
      color: '#c9404c'
    },
    {
      label: 'SWP',
      value: totals.SWP,
      color: '#2f9c75'
    },
    {
      label: 'Switch',
      value: totals.Switch,
      color: '#3f7edb'
    },
    {
      label: 'STP',
      value: totals.STP,
      color: '#8d65d8'
    }
  ]

  /* ---------------- UPLOAD ---------------- */

  async function uploadFile(e) {
    const file =
      e.target.files?.[0]

    if (!file) return

    setUploading(true)
    setMessage('Reading Excel file...')
    setError('')

    try {
      const buffer =
        await file.arrayBuffer()

      const workbook = XLSX.read(
        buffer,
        {
          type: 'array',
          cellDates: true
        }
      )

      const worksheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ]

      const raw =
        XLSX.utils.sheet_to_json(
          worksheet,
          { defval: null }
        )

      const mapped = raw
        .map(mapRow)
        .filter(
          row =>
            row.investor_name &&
            row.transaction_date &&
            row.amount != null
        )

      if (!mapped.length) {
        throw new Error(
          'No valid transactions found in the Excel file.'
        )
      }

      setMessage(
        `Uploading ${mapped.length} transactions...`
      )

      for (
        let i = 0;
        i < mapped.length;
        i += 500
      ) {
        const { error } =
          await supabase
            .from('transactions')
            .insert(
              mapped.slice(i, i + 500)
            )

        if (error) throw error
      }

      setMessage(
        'System is analysing and classifying transactions...'
      )

      /*
        This runs your SQL classification function.
        It should classify transactions as:
        Redemption / SWP / Switch / STP
      */

      const {
        error: rpcError
      } =
        await supabase.rpc(
          'run_redemption_classification'
        )

      if (rpcError) {
        throw rpcError
      }

      setMessage(
        'Done. Transactions analysed and dashboard updated.'
      )

      await loadData()
      await loadRMs()

    } catch (err) {
      setError(err.message)
      setMessage('')
    }

    setUploading(false)
    e.target.value = ''
  }

  /* ---------------- EXPORT ---------------- */

  function exportExcel() {
    const output = filtered.map(row => ({
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
        getSource(row),

      SystemClassification:
        getClassification(row)
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
      'Transaction Data'
    )

    XLSX.writeFile(
      workbook,
      'snowball-redemption-report.xlsx'
    )
  }

  function exportPDF() {
    const doc =
      new jsPDF({
        orientation: 'landscape'
      })

    doc.setFontSize(16)

    doc.text(
      'Snowball Redemption Tracker',
      14,
      14
    )

    doc.setFontSize(9)

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
          (row.scheme || '').slice(
            0,
            35
          ),
          money(row.amount),
          getSource(row),
          getClassification(row)
        ])
    })

    doc.save(
      'snowball-redemption-report.pdf'
    )
  }

  /* ---------------- RM MANAGEMENT ---------------- */

  async function addRM(e) {
    e.preventDefault()

    const name = newRM.trim()

    if (!name) return

    setRmMessage('')

    const { error } =
      await supabase
        .from('rms')
        .insert({
          name
        })

    if (error) {
      setRmMessage(
        error.message
      )
      return
    }

    setNewRM('')
    setRmMessage(
      'RM added successfully.'
    )

    await loadRMs()
  }

  async function deleteRM(item) {
    if (
      !window.confirm(
        `Delete ${item.name}?`
      )
    ) {
      return
    }

    if (item.fallback) {
      setRmMessage(
        'This RM exists in transaction data. Add the RM table SQL setup if you want to manage it separately.'
      )

      return
    }

    const { error } =
      await supabase
        .from('rms')
        .delete()
        .eq('id', item.id)

    if (error) {
      setRmMessage(
        error.message
      )
      return
    }

    await loadRMs()
  }

  /* -------------------------------------------------------
     LOGIN SCREEN
  ------------------------------------------------------- */

  if (!session) {
    return (
      <main className="loginPage">

        <section className="loginCard">

          <div className="loginLogo">
            <img
              src="/snowball-logo.png"
              alt="Snowball Financial Services"
              onError={e => {
                e.currentTarget.style.display =
                  'none'
              }}
            />

            <div className="loginLogoFallback">
              <strong>snowball</strong>
              <span>
                FINANCIAL SERVICES
              </span>
            </div>
          </div>

          <h1>
            Snowball Redemption Tracker
          </h1>

          {!resetMode ? (
            <>
              <p>
                Sign in to access your
                transaction dashboard
              </p>

              <form onSubmit={login}>

                <input
                  type="email"
                  placeholder="Email address"
                  value={email}
                  onChange={e =>
                    setEmail(
                      e.target.value
                    )
                  }
                  required
                />

                <input
                  type="password"
                  placeholder="Password"
                  value={password}
                  onChange={e =>
                    setPassword(
                      e.target.value
                    )
                  }
                  required
                />

                <button
                  className="loginButton"
                  disabled={loading}
                >
                  {loading
                    ? 'Signing in...'
                    : 'Login'}
                </button>

              </form>

              <button
                className="forgotButton"
                onClick={() => {
                  setResetMode(true)
                  setError('')
                }}
              >
                Forgot / Reset Password?
              </button>
            </>
          ) : (
            <>
              <p>
                Enter your email address.
                We will send you a password
                reset link.
              </p>

              <form
                onSubmit={
                  sendPasswordReset
                }
              >
                <input
                  type="email"
                  placeholder="Email address"
                  value={email}
                  onChange={e =>
                    setEmail(
                      e.target.value
                    )
                  }
                  required
                />

                <button
                  className="loginButton"
                  disabled={loading}
                >
                  {loading
                    ? 'Sending...'
                    : 'Send Reset Link'}
                </button>
              </form>

              <button
                className="forgotButton"
                onClick={() => {
                  setResetMode(false)
                  setResetSent(false)
                  setError('')
                }}
              >
                <ArrowLeft size={15} />
                Back to Login
              </button>
            </>
          )}

          {error && (
            <div className="loginError">
              {error}
            </div>
          )}

          {resetSent && (
            <div className="successBox">
              Password reset link has been
              sent. Please check your email.
            </div>
          )}

        </section>

      </main>
    )
  }

  /* -------------------------------------------------------
     MAIN APPLICATION
  ------------------------------------------------------- */

  return (
    <div className="appShell">

      {/* SIDEBAR */}

      <aside className="sidebar">

        <div className="sidebarLogo">

          <img
            src="/snowball-logo.png"
            alt="Snowball Financial Services"
            onError={e => {
              e.currentTarget.style.display =
                'none'
            }}
          />

          <div className="sidebarLogoFallback">
            <strong>snowball</strong>
            <span>
              FINANCIAL SERVICES
            </span>
          </div>

        </div>

        <nav className="navMenu">

          <button
            className={
              activeTab === 'dashboard'
                ? 'navItem active'
                : 'navItem'
            }
            onClick={() =>
              setActiveTab('dashboard')
            }
          >
            <LayoutDashboard size={19} />
            Dashboard
          </button>

          <button
            className={
              activeTab ===
              'transactions'
                ? 'navItem active'
                : 'navItem'
            }
            onClick={() =>
              setActiveTab(
                'transactions'
              )
            }
          >
            <Table2 size={19} />
            Transaction Data
          </button>

          <button
            className={
              activeTab === 'rms'
                ? 'navItem active'
                : 'navItem'
            }
            onClick={() =>
              setActiveTab('rms')
            }
          >
            <Users size={19} />
            Manage RMs
          </button>

        </nav>

        <div className="sidebarFooter">
          Snowball Financial Services
        </div>

      </aside>

      {/* MAIN */}

      <main className="mainContent">

        <header className="topHeader">

          <div>
            <h1>
              Snowball Redemption Tracker
            </h1>

            <p>
              Analyse transactions and
              monitor redemption activity
            </p>
          </div>

          <div className="headerActions">

            <button
              className="refreshIcon"
              onClick={loadData}
              title="Refresh data"
            >
              <RefreshCw size={16} />
            </button>

            <button
              className="logoutButton"
              onClick={logout}
            >
              <LogOut size={16} />
              Logout
            </button>

          </div>

        </header>

        {error && (
          <div className="errorBanner">
            <X size={17} />
            {error}
          </div>
        )}

        {/* FILTER TOOLBAR */}

        {activeTab !== 'rms' && (
          <section className="toolbar">

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
              onChange={e =>
                setRm(
                  e.target.value
                )
              }
            >
              <option>All</option>

              {rms.map(item => (
                <option
                  key={
                    item.id ||
                    item.name
                  }
                  value={item.name}
                >
                  {item.name}
                </option>
              ))}
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

            <button
              className="uploadButton"
              disabled={uploading}
            >
              <Upload size={16} />

              <label>
                {uploading
                  ? 'Uploading...'
                  : 'Upload Excel'}

                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  onChange={uploadFile}
                />
              </label>
            </button>

            <button
              className="toolButton"
              onClick={exportExcel}
            >
              <Download size={16} />
              Excel
            </button>

            <button
              className="toolButton"
              onClick={exportPDF}
            >
              <FileText size={16} />
              PDF
            </button>

          </section>
        )}

        {message && (
          <div className="messageBanner">
            {message}
          </div>
        )}

        {/* ============================================
            DASHBOARD
        ============================================ */}

        {activeTab === 'dashboard' && (
          <>

            {/* TOP CARDS */}

            <section className="summaryCards">

              <article className="summaryCard redemption">
                <span>Redemption</span>
                <strong>
                  {money(
                    totals.Redemption
                  )}
                </strong>
              </article>

              <article className="summaryCard swp">
                <span>SWP</span>
                <strong>
                  {money(
                    totals.SWP
                  )}
                </strong>
              </article>

              <article className="summaryCard switch">
                <span>Switch</span>
                <strong>
                  {money(
                    totals.Switch
                  )}
                </strong>
              </article>

              <article className="summaryCard stp">
                <span>STP</span>
                <strong>
                  {money(
                    totals.STP
                  )}
                </strong>
              </article>

              <article className="summaryCard">
                <span>Investors</span>
                <strong>
                  {totals.Investors}
                </strong>
              </article>

              <article className="summaryCard">
                <span>Transactions</span>
                <strong>
                  {totals.Transactions}
                </strong>
              </article>

            </section>

            {/* ANALYTICS */}

            <section className="analyticsGrid">

              <article className="panel classificationPanel">

                <h2>
                  Amount by Classification
                </h2>

                <DonutChart
                  data={classificationData}
                />

              </article>

              <article className="panel trendPanel">

                <h2>
                  Monthly Trend (Amount in ₹)
                </h2>

                <MonthlyTrend
                  rows={filtered}
                />

              </article>

              <article className="panel analysisPanel">

                <h2>
                  Classification Analysis
                </h2>

                <div className="analysisBars">

                  {classificationData.map(
                    item => {
                      const max =
                        Math.max(
                          ...classificationData.map(
                            x => x.value
                          ),
                          1
                        )

                      return (
                        <div
                          className="analysisRow"
                          key={item.label}
                        >

                          <span>
                            {item.label}
                          </span>

                          <div className="analysisTrack">
                            <div
                              className="analysisFill"
                              style={{
                                width: `${
                                  (item.value /
                                    max) *
                                  100
                                }%`,
                                background:
                                  item.color
                              }}
                            />
                          </div>

                          <b>
                            {shortMoney(
                              item.value
                            )}
                          </b>

                        </div>
                      )
                    }
                  )}

                </div>

              </article>

              <article className="panel rmPanel">

                <h2>
                  Transactions by RM
                </h2>

                <p className="panelSub">
                  Total transaction value
                </p>

                <RMBarChart
                  rows={filtered}
                />

              </article>

              <article className="panel recentPanel">

                <div className="panelHeader">

                  <div>
                    <h2>
                      Recent Transactions
                    </h2>
                  </div>

                  <button
                    className="viewButton"
                    onClick={() =>
                      setActiveTab(
                        'transactions'
                      )
                    }
                  >
                    View All Transactions
                  </button>

                </div>

                <div className="miniTableWrap">

                  <table className="miniTable">

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
                        .slice(0, 6)
                        .map(row => (
                          <tr
                            key={row.id}
                          >
                            <td>
                              {row.transaction_date}
                            </td>

                            <td>
                              {row.rm_name}
                            </td>

                            <td>
                              {row.investor_name}
                            </td>

                            <td>
                              {row.scheme}
                            </td>

                            <td>
                              {money(
                                row.amount
                              )}
                            </td>

                            <td>
                              <span className="sourceBadge">
                                {getSource(
                                  row
                                )}
                              </span>
                            </td>

                            <td>
                              <span
                                className={`classificationBadge ${getClassification(row)
                                  .toLowerCase()
                                  .replace(
                                    /\s/g,
                                    ''
                                  )}`}
                              >
                                {getClassification(
                                  row
                                )}
                              </span>
                            </td>
                          </tr>
                        ))}

                    </tbody>

                  </table>

                </div>

              </article>

            </section>

          </>
        )}

        {/* ============================================
            TRANSACTION DATA
        ============================================ */}

        {activeTab === 'transactions' && (
          <section className="dataPanel">

            <div className="sectionTitle">
              <div>
                <h2>
                  Transaction Data
                </h2>

                <p>
                  Detailed transaction information
                  and system classification
                </p>
              </div>

              <strong>
                {filtered.length} Transactions
              </strong>
            </div>

            <div className="transactionTableWrap">

              <table className="transactionTable">

                <thead>
                  <tr>
                    <th>Date</th>
                    <th>RM</th>
                    <th>Investor</th>
                    <th>Scheme</th>
                    <th>Amount</th>
                    <th>Source</th>
                    <th>
                      System Classification
                    </th>
                  </tr>
                </thead>

                <tbody>

                  {filtered
                    .slice(0, 1000)
                    .map(row => (
                      <tr key={row.id}>

                        <td>
                          {row.transaction_date}
                        </td>

                        <td>
                          {row.rm_name}
                        </td>

                        <td>
                          {row.investor_name}
                        </td>

                        <td>
                          {row.scheme}
                        </td>

                        <td className="amountCell">
                          {money(
                            row.amount
                          )}
                        </td>

                        <td>
                          {getSource(row)}
                        </td>

                        <td>
                          <span
                            className={`classificationBadge ${getClassification(row)
                              .toLowerCase()
                              .replace(
                                /\s/g,
                                ''
                              )}`}
                          >
                            {getClassification(
                              row
                            )}
                          </span>
                        </td>

                      </tr>
                    ))}

                </tbody>

              </table>

            </div>

          </section>
        )}

        {/* ============================================
            MANAGE RMS
        ============================================ */}

        {activeTab === 'rms' && (
          <section className="manageRMPanel">

            <div className="sectionTitle">

              <div>
                <h2>
                  Manage RMs
                </h2>

                <p>
                  Add or remove Relationship
                  Managers as your team changes.
                </p>
              </div>

            </div>

            <form
              className="addRMForm"
              onSubmit={addRM}
            >

              <input
                placeholder="Enter RM name"
                value={newRM}
                onChange={e =>
                  setNewRM(
                    e.target.value
                  )
                }
              />

              <button
                className="addRMButton"
              >
                <Plus size={17} />
                Add RM
              </button>

            </form>

            {rmMessage && (
              <div className="rmMessage">
                {rmMessage}
              </div>
            )}

            <div className="rmList">

              {rms.length === 0 && (
                <div className="emptyState">
                  No RMs have been added yet.
                </div>
              )}

              {rms.map(item => (
                <div
                  className="rmListItem"
                  key={
                    item.id ||
                    item.name
                  }
                >

                  <div className="rmAvatar">
                    {item.name
                      ?.charAt(0)
                      ?.toUpperCase()}
                  </div>

                  <strong>
                    {item.name}
                  </strong>

                  <button
                    onClick={() =>
                      deleteRM(item)
                    }
                    title="Delete RM"
                  >
                    <Trash2 size={17} />
                  </button>

                </div>
              ))}

            </div>

          </section>
        )}

      </main>

    </div>
  )
}

createRoot(
  document.getElementById('root')
).render(<App />)
