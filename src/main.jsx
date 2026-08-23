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
  RefreshCw,
  LogOut,
  Plus,
  Trash2,
  KeyRound,
  ArrowUpRight,
  ChevronRight
} from 'lucide-react'

import './styles.css'
import logo from './logo.png'

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
)

/* -------------------- HELPERS -------------------- */

const money = (n) =>
  new Intl.NumberFormat('en-IN', {
    style: 'currency',
    currency: 'INR',
    maximumFractionDigits: 0
  }).format(Number(n || 0))

const shortMoney = (n) => {
  const value = Number(n || 0)

  if (value >= 10000000) return `₹${(value / 10000000).toFixed(2)} Cr`
  if (value >= 100000) return `₹${(value / 100000).toFixed(1)} L`
  if (value >= 1000) return `₹${(value / 1000).toFixed(1)} K`

  return money(value)
}

const iso = (value) => {
  if (!value) return null

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10)
  }

  const date = new Date(value)

  if (Number.isNaN(date.getTime())) return null

  return date.toISOString().slice(0, 10)
}

const norm = (value) =>
  String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')

const getClassification = (row) => {
  const existing =
    row.classified_transaction_type ||
    row.classification_status ||
    row.transaction_type ||
    ''

  const original = String(
    row.original_transaction_type || row.type || ''
  ).toLowerCase()

  if (original.includes('stp')) return 'STP'
  if (original.includes('switch')) return 'Switch'
  if (original.includes('swp')) return 'SWP'

  if (
    String(existing).toLowerCase().includes('swp')
  ) return 'SWP'

  if (
    String(existing).toLowerCase().includes('switch')
  ) return 'Switch'

  if (
    String(existing).toLowerCase().includes('stp')
  ) return 'STP'

  return 'Redemption'
}

function mapRow(row) {
  const lookup = Object.fromEntries(
    Object.entries(row).map(([key, value]) => [norm(key), value])
  )

  const get = (...keys) =>
    keys.map((key) => lookup[norm(key)]).find((value) => value !== undefined)

  const amountRaw = get('Amount(₹)', 'Amount', 'amount')

  const amount =
    typeof amountRaw === 'number'
      ? amountRaw
      : Number(String(amountRaw || '').replace(/[₹,\s]/g, ''))

  return {
    rm_name:
      get('Partner/Employee', 'RM', 'RM Name', 'rm_name') || null,

    group_name:
      get('Group', 'group_name') || null,

    investor_name:
      get('Investor', 'Investor Name', 'investor_name') || null,

    transaction_date:
      iso(get('Date', 'Transaction Date', 'transaction_date')),

    folio_no:
      String(
        get(
          'Folio No/Demat A/C',
          'Folio',
          'Folio No',
          'folio_no'
        ) || ''
      ) || null,

    scheme:
      get('Scheme', 'Fund', 'scheme') || null,

    amount:
      Number.isFinite(amount) ? amount : null,

    original_transaction_type:
      get('Type', 'Transaction Type', 'original_transaction_type') || null
  }
}

/* -------------------- APP -------------------- */

function App() {
  const [session, setSession] = useState(null)

  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)

  const [activePage, setActivePage] = useState('dashboard')

  const [rows, setRows] = useState([])
  const [rms, setRms] = useState([])

  const [selectedRM, setSelectedRM] = useState('All')
  const [period, setPeriod] = useState('YTD')
  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')

  const [uploading, setUploading] = useState(false)
  const [message, setMessage] = useState('')

  const [newRM, setNewRM] = useState('')

  const [forgotPassword, setForgotPassword] = useState(false)
  const [resetMessage, setResetMessage] = useState('')

  /* ---------------- AUTH ---------------- */

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session)
    })

    const {
      data: { subscription }
    } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(newSession)
    })

    return () => subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (session) {
      loadData()
      loadRMs()
    }
  }, [session])

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

  async function sendResetLink(event) {
    event.preventDefault()

    setLoading(true)
    setResetMessage('')
    setError('')

    const { error: resetError } =
      await supabase.auth.resetPasswordForEmail(email, {
        redirectTo: window.location.origin
      })

    if (resetError) {
      setError(resetError.message)
    } else {
      setResetMessage(
        'Password reset link has been sent. Please check your email.'
      )
    }

    setLoading(false)
  }

  /* ---------------- DATA ---------------- */

  async function loadData() {
    setLoading(true)

    const { data, error: dataError } = await supabase
      .from('transactions')
      .select('*')
      .order('transaction_date', { ascending: false })

    if (dataError) {
      setError(dataError.message)
      setLoading(false)
      return
    }

    setRows(data || [])

    const transactionRMs = Array.from(
      new Set(
        (data || [])
          .map((item) => item.rm_name)
          .filter(Boolean)
      )
    )

    setRms((current) =>
      Array.from(new Set([...current, ...transactionRMs])).sort()
    )

    setLoading(false)
  }

  async function loadRMs() {
    const { data, error: rmError } = await supabase
      .from('rms')
      .select('*')
      .order('name', { ascending: true })

    if (!rmError && data) {
      setRms(
        Array.from(
          new Set(data.map((item) => item.name).filter(Boolean))
        )
      )
    }
  }

  /* ---------------- FILTERING ---------------- */

  const filteredRows = useMemo(() => {
    return rows.filter((item) => {
      if (
        selectedRM !== 'All' &&
        item.rm_name !== selectedRM
      ) {
        return false
      }

      const date = item.transaction_date

      if (!date) return false

      if (from && date < from) return false
      if (to && date > to) return false

      if (!from && !to) {
        const now = new Date()
        const currentDate = new Date(`${date}T00:00:00`)

        if (period === 'WTD') {
          const day = (now.getDay() + 6) % 7
          const start = new Date(now)

          start.setDate(now.getDate() - day)
          start.setHours(0, 0, 0, 0)

          if (currentDate < start) return false
        }

        if (period === 'MTD') {
          if (
            currentDate.getMonth() !== now.getMonth() ||
            currentDate.getFullYear() !== now.getFullYear()
          ) {
            return false
          }
        }

        if (period === 'QTD') {
          const currentQuarter = Math.floor(now.getMonth() / 3)
          const transactionQuarter = Math.floor(
            currentDate.getMonth() / 3
          )

          if (
            currentDate.getFullYear() !== now.getFullYear() ||
            transactionQuarter !== currentQuarter
          ) {
            return false
          }
        }

        if (period === 'YTD') {
          if (
            currentDate.getFullYear() !== now.getFullYear()
          ) {
            return false
          }
        }
      }

      return true
    })
  }, [
    rows,
    selectedRM,
    period,
    from,
    to
  ])

  /* ---------------- TOTALS ---------------- */

  const totals = useMemo(() => {
    const result = {
      Redemption: 0,
      SWP: 0,
      Switch: 0,
      STP: 0,
      Investors: new Set(),
      Transactions: 0
    }

    filteredRows.forEach((item) => {
      const classification = getClassification(item)
      const amount = Number(item.amount || 0)

      if (classification === 'Redemption') {
        result.Redemption += amount
      }

      if (classification === 'SWP') {
        result.SWP += amount
      }

      if (classification === 'Switch') {
        result.Switch += amount
      }

      if (classification === 'STP') {
        result.STP += amount
      }

      if (item.investor_name) {
        result.Investors.add(item.investor_name)
      }

      result.Transactions += 1
    })

    return {
      ...result,
      Investors: result.Investors.size
    }
  }, [filteredRows])

  /* ---------------- RM ANALYSIS ---------------- */

  const rmData = useMemo(() => {
    const map = {}

    filteredRows.forEach((item) => {
      const name = item.rm_name || 'Unassigned'

      if (!map[name]) {
        map[name] = 0
      }

      map[name] += Number(item.amount || 0)
    })

    return Object.entries(map)
      .map(([name, value]) => ({
        name,
        value
      }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 8)
  }, [filteredRows])

  /* ---------------- MONTHLY DATA ---------------- */

  const monthlyData = useMemo(() => {
    const months = {}

    filteredRows.forEach((item) => {
      if (!item.transaction_date) return

      const date = new Date(
        `${item.transaction_date}T00:00:00`
      )

      const key = `${date.getFullYear()}-${String(
        date.getMonth() + 1
      ).padStart(2, '0')}`

      if (!months[key]) {
        months[key] = {
          Redemption: 0,
          SWP: 0,
          Switch: 0,
          STP: 0
        }
      }

      const classification = getClassification(item)

      if (months[key][classification] !== undefined) {
        months[key][classification] += Number(item.amount || 0)
      }
    })

    return Object.entries(months)
      .sort(([a], [b]) => a.localeCompare(b))
      .slice(-8)
      .map(([month, values]) => ({
        month: month.slice(5),
        ...values
      }))
  }, [filteredRows])

  /* ---------------- UPLOAD ---------------- */

  async function uploadFile(event) {
    const file = event.target.files?.[0]

    if (!file) return

    setUploading(true)
    setMessage('Reading Excel file...')
    setError('')

    try {
      const buffer = await file.arrayBuffer()

      const workbook = XLSX.read(buffer, {
        type: 'array',
        cellDates: true
      })

      const worksheet =
        workbook.Sheets[workbook.SheetNames[0]]

      const raw =
        XLSX.utils.sheet_to_json(worksheet, {
          defval: null
        })

      const mapped = raw
        .map(mapRow)
        .filter(
          (item) =>
            item.investor_name &&
            item.transaction_date &&
            item.amount !== null
        )

      if (!mapped.length) {
        throw new Error(
          'No valid transactions found in the uploaded Excel file.'
        )
      }

      setMessage(
        `Uploading ${mapped.length} transactions...`
      )

      for (let index = 0; index < mapped.length; index += 500) {
        const batch = mapped.slice(index, index + 500)

        const { error: uploadError } =
          await supabase
            .from('transactions')
            .insert(batch)

        if (uploadError) {
          throw uploadError
        }
      }

      /*
       * Run database classification if the function exists.
       * If not, dashboard still uses available classifications.
       */
      const { error: rpcError } =
        await supabase.rpc(
          'run_redemption_classification'
        )

      if (rpcError) {
        console.warn(
          'Classification function warning:',
          rpcError.message
        )
      }

      setMessage(
        'Upload completed successfully. Dashboard updated.'
      )

      await loadData()
    } catch (uploadError) {
      setError(uploadError.message)
      setMessage('')
    }

    setUploading(false)
    event.target.value = ''
  }

  /* ---------------- EXPORTS ---------------- */

  function exportExcel() {
    const output = filteredRows.map((item) => ({
      Date: item.transaction_date,
      RM: item.rm_name,
      Investor: item.investor_name,
      Folio: item.folio_no,
      Scheme: item.scheme,
      Amount: item.amount,
      Source: getClassification(item),
      Classification: getClassification(item)
    }))

    const worksheet =
      XLSX.utils.json_to_sheet(output)

    const workbook =
      XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'Snowball Transactions'
    )

    XLSX.writeFile(
      workbook,
      'snowball-redemption-report.xlsx'
    )
  }

  function exportPDF() {
    const document = new jsPDF({
      orientation: 'landscape'
    })

    document.setFontSize(16)
    document.text(
      'Snowball Redemption Tracker',
      14,
      14
    )

    document.setFontSize(10)

    document.text(
      `RM: ${selectedRM}`,
      14,
      21
    )

    autoTable(document, {
      startY: 27,

      head: [[
        'Date',
        'RM',
        'Investor',
        'Scheme',
        'Amount',
        'Classification'
      ]],

      body: filteredRows.map((item) => [
        item.transaction_date,
        item.rm_name,
        item.investor_name,
        String(item.scheme || '').slice(0, 35),
        money(item.amount),
        getClassification(item)
      ])
    })

    document.save(
      'snowball-redemption-report.pdf'
    )
  }

  /* ---------------- RM MANAGEMENT ---------------- */

  async function addRM() {
    const name = newRM.trim()

    if (!name) return

    setError('')

    const { error: rmError } =
      await supabase
        .from('rms')
        .insert([{ name }])

    if (rmError) {
      /*
       * If rms table does not exist yet,
       * still allow temporary local display.
       */
      if (
        !rms.some(
          (item) =>
            item.toLowerCase() === name.toLowerCase()
        )
      ) {
        setRms([...rms, name].sort())
      }

      setMessage(
        'RM added locally. Please create the rms table in Supabase to save it permanently.'
      )
    } else {
      setMessage('RM added successfully.')
      await loadRMs()
    }

    setNewRM('')
  }

  async function deleteRM(name) {
    const confirmed = window.confirm(
      `Delete ${name} from the RM list?`
    )

    if (!confirmed) return

    const { error: rmError } =
      await supabase
        .from('rms')
        .delete()
        .eq('name', name)

    if (rmError) {
      setRms(
        rms.filter((item) => item !== name)
      )

      setMessage(
        'RM removed from the current list.'
      )
    } else {
      setMessage('RM removed successfully.')
      await loadRMs()
    }

    if (selectedRM === name) {
      setSelectedRM('All')
    }
  }

  /* ---------------- LOGIN SCREEN ---------------- */

  if (!session) {
    return (
      <main className="loginPage">
        <section className="loginCard">

          <img
            src={logo}
            alt="Snowball Financial Services"
            className="loginLogo"
          />

          <div className="loginHeading">
            <h1>Redemption Tracker</h1>
            <p>
              Secure transaction monitoring dashboard
            </p>
          </div>

          {!forgotPassword ? (
            <form onSubmit={login}>
              <input
                type="email"
                placeholder="Email"
                value={email}
                onChange={(event) =>
                  setEmail(event.target.value)
                }
                required
              />

              <input
                type="password"
                placeholder="Password"
                value={password}
                onChange={(event) =>
                  setPassword(event.target.value)
                }
                required
              />

              <button
                type="submit"
                className="primaryButton"
                disabled={loading}
              >
                {loading
                  ? 'Signing in...'
                  : 'Login'}
              </button>

              <button
                type="button"
                className="textButton"
                onClick={() => {
                  setForgotPassword(true)
                  setError('')
                  setResetMessage('')
                }}
              >
                Forgot / Reset Password?
              </button>

              {error && (
                <div className="formError">
                  {error}
                </div>
              )}
            </form>
          ) : (
            <form onSubmit={sendResetLink}>
              <p className="resetText">
                Enter your registered email address.
              </p>

              <input
                type="email"
                placeholder="Email"
                value={email}
                onChange={(event) =>
                  setEmail(event.target.value)
                }
                required
              />

              <button
                type="submit"
                className="primaryButton"
                disabled={loading}
              >
                {loading
                  ? 'Sending...'
                  : 'Send Reset Link'}
              </button>

              <button
                type="button"
                className="textButton"
                onClick={() => {
                  setForgotPassword(false)
                  setError('')
                }}
              >
                Back to Login
              </button>

              {resetMessage && (
                <div className="successMessage">
                  {resetMessage}
                </div>
              )}

              {error && (
                <div className="formError">
                  {error}
                </div>
              )}
            </form>
          )}
        </section>
      </main>
    )
  }

  /* ---------------- DASHBOARD ---------------- */

  const maxRMValue =
    Math.max(
      ...rmData.map((item) => item.value),
      1
    )

  const classificationItems = [
    {
      label: 'Redemption',
      value: totals.Redemption,
      color: 'redemption'
    },
    {
      label: 'SWP',
      value: totals.SWP,
      color: 'swp'
    },
    {
      label: 'Switch',
      value: totals.Switch,
      color: 'switch'
    },
    {
      label: 'STP',
      value: totals.STP,
      color: 'stp'
    }
  ]

  return (
    <div className="appShell">

      {/* SIDEBAR */}

      <aside className="sidebar">

        <div className="brand">
          <img
            src={logo}
            alt="Snowball Financial Services"
          />
        </div>

        <nav className="sideNav">

          <button
            className={
              activePage === 'dashboard'
                ? 'navItem active'
                : 'navItem'
            }
            onClick={() =>
              setActivePage('dashboard')
            }
          >
            <LayoutDashboard size={18} />
            Dashboard
          </button>

          <button
            className={
              activePage === 'transactions'
                ? 'navItem active'
                : 'navItem'
            }
            onClick={() =>
              setActivePage('transactions')
            }
          >
            <Table2 size={18} />
            Transaction Data
          </button>

          <button
            className={
              activePage === 'rms'
                ? 'navItem active'
                : 'navItem'
            }
            onClick={() =>
              setActivePage('rms')
            }
          >
            <Users size={18} />
            Manage RMs
          </button>

        </nav>

        <div className="sidebarFooter">
          <span>Snowball Financial Services</span>
        </div>

      </aside>

      {/* MAIN */}

      <main className="mainArea">

        <header className="topHeader">
          <div>
            <h1>
              {activePage === 'dashboard' &&
                'Snowball Redemption Tracker'}

              {activePage === 'transactions' &&
                'Transaction Data'}

              {activePage === 'rms' &&
                'Manage Relationship Managers'}
            </h1>

            <p>
              Analyse transactions and monitor redemption activity
            </p>
          </div>

          <button
            className="logoutButton"
            onClick={() =>
              supabase.auth.signOut()
            }
          >
            <LogOut size={16} />
            Logout
          </button>
        </header>

        {error && (
          <div className="errorBanner">
            {error}
          </div>
        )}

        {message && (
          <div className="successBanner">
            {message}
          </div>
        )}

        {/* CONTROLS */}

        {activePage !== 'rms' && (
          <section className="controlPanel">

            <div className="periodButtons">
              {['WTD', 'MTD', 'QTD', 'YTD'].map(
                (item) => (
                  <button
                    key={item}
                    className={
                      period === item &&
                      !from &&
                      !to
                        ? 'periodButton selected'
                        : 'periodButton'
                    }
                    onClick={() => {
                      setPeriod(item)
                      setFrom('')
                      setTo('')
                    }}
                  >
                    {item}
                  </button>
                )
              )}
            </div>

            <select
              value={selectedRM}
              onChange={(event) =>
                setSelectedRM(event.target.value)
              }
            >
              <option>All</option>

              {rms.map((item) => (
                <option key={item}>
                  {item}
                </option>
              ))}
            </select>

            <input
              type="date"
              value={from}
              onChange={(event) =>
                setFrom(event.target.value)
              }
            />

            <input
              type="date"
              value={to}
              onChange={(event) =>
                setTo(event.target.value)
              }
            />

            <label
              className={
                uploading
                  ? 'uploadButton disabled'
                  : 'uploadButton'
              }
            >
              <Upload size={16} />

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

            <button
              className="actionButton"
              onClick={exportExcel}
            >
              <Download size={16} />
              Excel
            </button>

            <button
              className="actionButton"
              onClick={exportPDF}
            >
              <FileText size={16} />
              PDF
            </button>

            <button
              className="iconButton"
              onClick={loadData}
              title="Refresh"
            >
              <RefreshCw size={17} />
            </button>

          </section>
        )}

        {/* DASHBOARD PAGE */}

        {activePage === 'dashboard' && (
          <>
            <section className="metricGrid">

              <MetricCard
                title="Redemption"
                value={money(totals.Redemption)}
                type="redemption"
              />

              <MetricCard
                title="SWP"
                value={money(totals.SWP)}
                type="swp"
              />

              <MetricCard
                title="Switch"
                value={money(totals.Switch)}
                type="switch"
              />

              <MetricCard
                title="STP"
                value={money(totals.STP)}
                type="stp"
              />

              <MetricCard
                title="Investors"
                value={totals.Investors}
                type="investors"
              />

              <MetricCard
                title="Transactions"
                value={totals.Transactions}
                type="transactions"
              />

            </section>

            <section className="dashboardGrid">

              <div className="panel classificationPanel">
                <h2>Amount by Classification</h2>

                <div className="classificationContent">

                  <div
                    className="donut"
                    style={{
                      background:
                        `conic-gradient(
                          #d85b61 0deg ${
                            totals.Transactions
                              ? (totals.Redemption /
                                  (totals.Redemption +
                                    totals.SWP +
                                    totals.Switch +
                                    totals.STP || 1)) *
                                360
                              : 0
                          }deg,
                          #54a77e ${
                            totals.Transactions
                              ? (totals.Redemption /
                                  (totals.Redemption +
                                    totals.SWP +
                                    totals.Switch +
                                    totals.STP || 1)) *
                                360
                              : 0
                          }deg ${
                            totals.Transactions
                              ? ((totals.Redemption +
                                  totals.SWP) /
                                  (totals.Redemption +
                                    totals.SWP +
                                    totals.Switch +
                                    totals.STP || 1)) *
                                360
                              : 0
                          }deg,
                          #668bc7 ${
                            ((totals.Redemption +
                              totals.SWP) /
                              (totals.Redemption +
                                totals.SWP +
                                totals.Switch +
                                totals.STP || 1)) *
                            360
                          }deg ${
                            ((totals.Redemption +
                              totals.SWP +
                              totals.Switch) /
                              (totals.Redemption +
                                totals.SWP +
                                totals.Switch +
                                totals.STP || 1)) *
                            360
                          }deg,
                          #9a75c9 ${
                            ((totals.Redemption +
                              totals.SWP +
                              totals.Switch) /
                              (totals.Redemption +
                                totals.SWP +
                                totals.Switch +
                                totals.STP || 1)) *
                            360
                          }deg 360deg
                        )`
                    }}
                  >
                    <div className="donutHole">
                      {totals.Transactions}
                    </div>
                  </div>

                  <div className="legend">
                    {classificationItems.map(
                      (item) => (
                        <div
                          className="legendItem"
                          key={item.label}
                        >
                          <span
                            className={`legendDot ${item.color}`}
                          />

                          <span>{item.label}</span>

                          <strong>
                            {shortMoney(item.value)}
                          </strong>
                        </div>
                      )
                    )}
                  </div>

                </div>
              </div>

              <div className="panel monthlyPanel">
                <h2>Monthly Trend (Amount in ₹)</h2>

                <div className="trendChart">
                  {monthlyData.length ? (
                    monthlyData.map((item) => {
                      const total =
                        item.Redemption +
                        item.SWP +
                        item.Switch +
                        item.STP

                      const max = Math.max(
                        ...monthlyData.map(
                          (month) =>
                            month.Redemption +
                            month.SWP +
                            month.Switch +
                            month.STP
                        ),
                        1
                      )

                      return (
                        <div
                          className="trendColumn"
                          key={item.month}
                        >
                          <div className="trendBars">

                            <span
                              className="trendBar redemption"
                              style={{
                                height: `${Math.max(
                                  4,
                                  (item.Redemption / max) *
                                    130
                                )}px`
                              }}
                            />

                            <span
                              className="trendBar swp"
                              style={{
                                height: `${Math.max(
                                  3,
                                  (item.SWP / max) * 130
                                )}px`
                              }}
                            />

                            <span
                              className="trendBar switch"
                              style={{
                                height: `${Math.max(
                                  3,
                                  (item.Switch / max) *
                                    130
                                )}px`
                              }}
                            />

                            <span
                              className="trendBar stp"
                              style={{
                                height: `${Math.max(
                                  3,
                                  (item.STP / max) * 130
                                )}px`
                              }}
                            />

                          </div>

                          <small>
                            {item.month}
                          </small>
                        </div>
                      )
                    })
                  ) : (
                    <div className="emptyChart">
                      Upload transaction data to view trend
                    </div>
                  )}
                </div>

                <div className="chartLegend">
                  <span>
                    <i className="redemption" />
                    Redemption
                  </span>

                  <span>
                    <i className="swp" />
                    SWP
                  </span>

                  <span>
                    <i className="switch" />
                    Switch
                  </span>

                  <span>
                    <i className="stp" />
                    STP
                  </span>
                </div>

              </div>

              <div className="panel analysisPanel">
                <h2>Classification Analysis</h2>

                <div className="analysisBars">
                  {classificationItems.map(
                    (item) => {
                      const total =
                        totals.Redemption +
                        totals.SWP +
                        totals.Switch +
                        totals.STP

                      const percentage =
                        total > 0
                          ? (item.value / total) * 100
                          : 0

                      return (
                        <div
                          className="analysisRow"
                          key={item.label}
                        >
                          <div className="analysisLabel">
                            <span
                              className={`legendDot ${item.color}`}
                            />

                            {item.label}
                          </div>

                          <div className="progressTrack">
                            <div
                              className={`progressFill ${item.color}`}
                              style={{
                                width: `${percentage}%`
                              }}
                            />
                          </div>

                          <strong>
                            {shortMoney(item.value)}
                          </strong>
                        </div>
                      )
                    }
                  )}
                </div>

              </div>

              <div className="panel rmPanel">
                <h2>Transactions by RM</h2>

                <div className="rmBars">
                  {rmData.length ? (
                    rmData.map((item) => (
                      <div
                        className="rmBarRow"
                        key={item.name}
                      >
                        <span>
                          {item.name}
                        </span>

                        <div className="rmTrack">
                          <div
                            className="rmFill"
                            style={{
                              width: `${
                                (item.value / maxRMValue) *
                                100
                              }%`
                            }}
                          />
                        </div>

                        <strong>
                          {shortMoney(item.value)}
                        </strong>
                      </div>
                    ))
                  ) : (
                    <div className="emptyChart">
                      No RM data available
                    </div>
                  )}
                </div>

              </div>

              <div className="panel recentPanel">
                <div className="panelHeader">
                  <h2>Recent Transactions</h2>

                  <button
                    className="viewAllButton"
                    onClick={() =>
                      setActivePage('transactions')
                    }
                  >
                    View All Transactions
                    <ChevronRight size={15} />
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
                        <th>Classification</th>
                      </tr>
                    </thead>

                    <tbody>
                      {filteredRows
                        .slice(0, 6)
                        .map((item) => (
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
                              {item.scheme}
                            </td>

                            <td>
                              {money(item.amount)}
                            </td>

                            <td>
                              <span
                                className={`classificationBadge ${getClassification(item).toLowerCase()}`}
                              >
                                {getClassification(item)}
                              </span>
                            </td>
                          </tr>
                        ))}
                    </tbody>
                  </table>
                </div>

              </div>

            </section>
          </>
        )}

        {/* TRANSACTION PAGE */}

        {activePage === 'transactions' && (
          <section className="dataPanel">

            <div className="panelHeader">
              <div>
                <h2>
                  Transaction Details
                </h2>

                <p>
                  {filteredRows.length} transactions found
                </p>
              </div>
            </div>

            <div className="fullTableWrap">

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
                  {filteredRows.map((item) => (
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
                        {item.scheme}
                      </td>

                      <td>
                        {money(item.amount)}
                      </td>

                      {/* Source now only shows Redemption / SWP / Switch / STP */}
                      <td>
                        {getClassification(item)}
                      </td>

                      <td>
                        <span
                          className={`classificationBadge ${getClassification(item).toLowerCase()}`}
                        >
                          {getClassification(item)}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>

              </table>

            </div>

          </section>
        )}

        {/* MANAGE RMS */}

        {activePage === 'rms' && (
          <section className="rmManagePanel">

            <div className="addRMBox">

              <div>
                <h2>
                  Add Relationship Manager
                </h2>

                <p>
                  Add or remove RMs as your organisation changes.
                </p>
              </div>

              <div className="addRMForm">
                <input
                  placeholder="Enter RM name"
                  value={newRM}
                  onChange={(event) =>
                    setNewRM(event.target.value)
                  }
                />

                <button
                  className="primaryButton"
                  onClick={addRM}
                >
                  <Plus size={17} />
                  Add RM
                </button>
              </div>

            </div>

            <div className="rmList">

              {rms.map((item) => (
                <div
                  className="rmListItem"
                  key={item}
                >
                  <div>
                    <Users size={18} />
                    {item}
                  </div>

                  <button
                    className="deleteButton"
                    onClick={() =>
                      deleteRM(item)
                    }
                  >
                    <Trash2 size={16} />
                    Delete
                  </button>
                </div>
              ))}

              {!rms.length && (
                <div className="emptyState">
                  No RMs added yet.
                </div>
              )}

            </div>

          </section>
        )}

      </main>
    </div>
  )
}

/* ---------------- METRIC CARD ---------------- */

function MetricCard({ title, value, type }) {
  return (
    <article className={`metricCard ${type}`}>
      <span>{title}</span>
      <strong>{value}</strong>
    </article>
  )
}

/* ---------------- RENDER ---------------- */

const rootElement =
  document.getElementById('root')

if (rootElement) {
  createRoot(rootElement).render(<App />)
}
