import React, { useEffect, useMemo, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'
import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'
import {
  LayoutDashboard,
  TableProperties,
  Users,
  Settings,
  LogOut,
  Upload,
  Download,
  FileText,
  RefreshCw,
  Plus,
  UserCheck,
  UserX,
  KeyRound,
  Mail,
  X
} from 'lucide-react'

import logo from './logo.png'
import './styles.css'

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

const shortMoney = n => {
  const value = Number(n || 0)

  if (value >= 10000000) {
    return `₹${(value / 10000000).toFixed(2)} Cr`
  }

  if (value >= 100000) {
    return `₹${(value / 100000).toFixed(2)} L`
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

function sourceType(value) {
  const v = String(value || '').toLowerCase()

  if (v.includes('swp')) return 'SWP'
  if (v.includes('switch')) return 'Switch'
  if (v.includes('stp')) return 'STP'
  return 'Redemption'
}

function mapRow(row) {
  const lookup = Object.fromEntries(
    Object.entries(row).map(([key, value]) => [norm(key), value])
  )

  const get = (...keys) =>
    keys.map(key => lookup[norm(key)]).find(value => value !== undefined)

  const amountRaw = get('Amount(₹)', 'Amount', 'amount')

  const amount =
    typeof amountRaw === 'number'
      ? amountRaw
      : Number(String(amountRaw || '').replace(/[₹,\s]/g, ''))

  const originalType =
    get(
      'Type',
      'Transaction Type',
      'transaction_type',
      'original_transaction_type'
    ) || 'Redemption'

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
        get('Folio No/Demat A/C', 'Folio', 'Folio No', 'folio_no') || ''
      ) || null,

    scheme:
      get('Scheme', 'Fund', 'scheme') || null,

    amount:
      Number.isFinite(amount) ? amount : null,

    transaction_type:
      'Imported',

    original_transaction_type:
      sourceType(originalType),

    classified_transaction_type:
      null,

    classification_reason:
      null
  }
}

function App() {
  const [session, setSession] = useState(null)

  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [loginError, setLoginError] = useState('')
  const [loading, setLoading] = useState(false)

  const [page, setPage] = useState('dashboard')

  const [rows, setRows] = useState([])
  const [rms, setRms] = useState([])
  const [admins, setAdmins] = useState([])

  const [isAdmin, setIsAdmin] = useState(false)

  const [rm, setRm] = useState('All')
  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')
  const [period, setPeriod] = useState('YTD')

  const [uploading, setUploading] = useState(false)
  const [message, setMessage] = useState('')
  const [error, setError] = useState('')

  const [newRM, setNewRM] = useState('')
  const [newAdmin, setNewAdmin] = useState('')

  const [forgotMode, setForgotMode] = useState(false)
  const [forgotEmail, setForgotEmail] = useState('')
  const [forgotMessage, setForgotMessage] = useState('')

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session)
    })

    const {
      data: { subscription }
    } = supabase.auth.onAuthStateChange((_event, currentSession) => {
      setSession(currentSession)
    })

    return () => subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (session) {
      initialise()
    }
  }, [session])

  async function initialise() {
    setError('')

    await Promise.all([
      loadData(),
      loadRMs(),
      checkAdmin()
    ])
  }

  async function checkAdmin() {
    const { data, error: adminError } =
      await supabase.rpc('is_app_admin')

    if (!adminError) {
      setIsAdmin(Boolean(data))

      if (data) {
        loadAdmins()
      }
    }
  }

  async function login(event) {
    event.preventDefault()

    setLoading(true)
    setLoginError('')

    const { error: authError } =
      await supabase.auth.signInWithPassword({
        email,
        password
      })

    if (authError) {
      setLoginError(authError.message)
    }

    setLoading(false)
  }

  async function logout() {
    setLoading(true)

    await supabase.auth.signOut()

    setSession(null)
    setRows([])
    setRms([])
    setAdmins([])
    setIsAdmin(false)
    setPage('dashboard')
    setRm('All')

    setLoading(false)
  }

  async function resetPassword(event) {
    event.preventDefault()

    setForgotMessage('')
    setLoginError('')

    if (!forgotEmail) {
      setForgotMessage('Please enter your email address.')
      return
    }

    const { error: resetError } =
      await supabase.auth.resetPasswordForEmail(forgotEmail, {
        redirectTo: window.location.origin
      })

    if (resetError) {
      setForgotMessage(resetError.message)
    } else {
      setForgotMessage(
        'Password reset link has been sent. Please check your email.'
      )
    }
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

  async function loadRMs() {
    const { data, error: rmError } =
      await supabase
        .from('relationship_managers')
        .select('*')
        .order('name')

    if (rmError) {
      console.error(rmError)
      return
    }

    setRms(data || [])
  }

  async function loadAdmins() {
    const { data, error: adminError } =
      await supabase
        .from('app_admins')
        .select('*')
        .order('email')

    if (!adminError) {
      setAdmins(data || [])
    }
  }

  const activeRMs = useMemo(() => {
    return rms
      .filter(item => item.is_active)
      .map(item => item.name)
  }, [rms])

  const filtered = useMemo(() => {
    return rows.filter(item => {
      if (rm !== 'All' && item.rm_name !== rm) {
        return false
      }

      const date = item.transaction_date

      if (!date) return false

      if (from && date < from) return false
      if (to && date > to) return false

      if (!from && !to) {
        const now = new Date()
        const dt = new Date(`${date}T00:00:00`)

        if (period === 'WTD') {
          const day = (now.getDay() + 6) % 7

          const start = new Date(now)

          start.setDate(now.getDate() - day)
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
          const currentQuarter =
            Math.floor(now.getMonth() / 3)

          if (
            dt.getFullYear() !== now.getFullYear() ||
            Math.floor(dt.getMonth() / 3) !== currentQuarter
          ) {
            return false
          }
        }

        if (
          period === 'YTD' &&
          dt.getFullYear() !== now.getFullYear()
        ) {
          return false
        }
      }

      return true
    })
  }, [rows, rm, from, to, period])

  const totals = useMemo(() => {
    const sum = type =>
      filtered
        .filter(
          item =>
            item.classified_transaction_type === type
        )
        .reduce(
          (total, item) =>
            total + Number(item.amount || 0),
          0
        )

    return {
      Redemption: sum('Redemption'),
      SWP: sum('SWP'),
      Switch: sum('Switch'),
      STP: sum('STP'),
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
        value: totals.Redemption,
        color: '#d45a61'
      },
      {
        name: 'SWP',
        value: totals.SWP,
        color: '#5c9b85'
      },
      {
        name: 'Switch',
        value: totals.Switch,
        color: '#5d7fa8'
      },
      {
        name: 'STP',
        value: totals.STP,
        color: '#8068a8'
      }
    ]
  }, [totals])

  const totalClassification = useMemo(() => {
    return classificationData.reduce(
      (sum, item) => sum + item.value,
      0
    )
  }, [classificationData])

  const donutBackground = useMemo(() => {
    if (!totalClassification) {
      return 'conic-gradient(#e8edf2 0deg 360deg)'
    }

    let start = 0

    const parts = classificationData.map(item => {
      const angle =
        (item.value / totalClassification) * 360

      const end = start + angle

      const result =
        `${item.color} ${start}deg ${end}deg`

      start = end

      return result
    })

    return `conic-gradient(${parts.join(', ')})`
  }, [classificationData, totalClassification])

  const rmData = useMemo(() => {
    const map = {}

    filtered.forEach(item => {
      const name = item.rm_name || 'Unassigned'

      map[name] =
        (map[name] || 0) +
        Number(item.amount || 0)
    })

    return Object.entries(map)
      .map(([name, value]) => ({
        name,
        value
      }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 7)
  }, [filtered])

  const monthlyData = useMemo(() => {
    const map = {}

    filtered.forEach(item => {
      if (!item.transaction_date) return

      const month =
        item.transaction_date.slice(0, 7)

      if (!map[month]) {
        map[month] = {
          Redemption: 0,
          SWP: 0,
          Switch: 0,
          STP: 0
        }
      }

      const type =
        item.classified_transaction_type ||
        'Redemption'

      if (map[month][type] !== undefined) {
        map[month][type] +=
          Number(item.amount || 0)
      }
    })

    return Object.entries(map)
      .sort(([a], [b]) => a.localeCompare(b))
      .slice(-8)
      .map(([month, values]) => ({
        month,
        ...values
      }))
  }, [filtered])

  async function uploadFile(event) {
    const file = event.target.files?.[0]

    if (!file) return

    setUploading(true)
    setError('')
    setMessage('Reading Excel file...')

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
          'No valid transactions found in the uploaded Excel file.'
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
        const { error: insertError } =
          await supabase
            .from('transactions')
            .insert(mapped.slice(i, i + 500))

        if (insertError) throw insertError
      }

      setMessage(
        'System is analysing SWP, Redemption, Switch and STP patterns...'
      )

      const { error: classifyError } =
        await supabase.rpc(
          'run_redemption_classification'
        )

      if (classifyError) throw classifyError

      setMessage(
        'Upload complete. Transactions have been classified.'
      )

      await loadData()
      await loadRMs()

      setTimeout(() => setMessage(''), 3500)

    } catch (uploadError) {
      setError(uploadError.message)
      setMessage('')
    }

    setUploading(false)
    event.target.value = ''
  }

  async function addRM(event) {
    event.preventDefault()

    const name = newRM.trim()

    if (!name) return

    const { error: rmError } =
      await supabase
        .from('relationship_managers')
        .insert({
          name,
          is_active: true
        })

    if (rmError) {
      setError(rmError.message)
      return
    }

    setNewRM('')
    await loadRMs()
  }

  async function toggleRM(item) {
    const { error: rmError } =
      await supabase
        .from('relationship_managers')
        .update({
          is_active: !item.is_active,
          updated_at: new Date().toISOString()
        })
        .eq('id', item.id)

    if (rmError) {
      setError(rmError.message)
      return
    }

    await loadRMs()
  }

  async function addAdmin(event) {
    event.preventDefault()

    const adminEmail =
      newAdmin.trim().toLowerCase()

    if (!adminEmail) return

    const { error: adminError } =
      await supabase
        .from('app_admins')
        .insert({
          email: adminEmail,
          is_active: true
        })

    if (adminError) {
      setError(adminError.message)
      return
    }

    setNewAdmin('')
    await loadAdmins()
  }

  async function toggleAdmin(item) {
    const { error: adminError } =
      await supabase
        .from('app_admins')
        .update({
          is_active: !item.is_active
        })
        .eq('id', item.id)

    if (adminError) {
      setError(adminError.message)
      return
    }

    await loadAdmins()
  }

  function exportExcel() {
    const output = filtered.map(item => ({
      Date: item.transaction_date,
      RM: item.rm_name,
      Investor: item.investor_name,
      Folio: item.folio_no,
      Scheme: item.scheme,
      Amount: item.amount,
      Source: item.original_transaction_type,
      Classification:
        item.classified_transaction_type
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
    const document = new jsPDF({
      orientation: 'landscape'
    })

    document.setFontSize(16)

    document.text(
      'Snowball Redemption Tracker',
      14,
      15
    )

    document.setFontSize(9)

    document.text(
      `RM: ${rm}`,
      14,
      22
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
        item.scheme,
        money(item.amount),
        item.original_transaction_type,
        item.classified_transaction_type
      ])
    })

    document.save(
      'snowball-transaction-report.pdf'
    )
  }

  function clearDates() {
    setFrom('')
    setTo('')
  }

  if (!session) {
    return (
      <main className="loginPage">
        <section className="loginCard">

          {!forgotMode ? (
            <>
              <h1>Snowball Redemption Tracker</h1>

              <p>
                Sign in to monitor redemption activity
              </p>

              <form onSubmit={login}>
                <input
                  type="email"
                  placeholder="Email address"
                  value={email}
                  onChange={event =>
                    setEmail(event.target.value)
                  }
                  required
                />

                <input
                  type="password"
                  placeholder="Password"
                  value={password}
                  onChange={event =>
                    setPassword(event.target.value)
                  }
                  required
                />

                <button
                  className="primaryLogin"
                  disabled={loading}
                >
                  {loading
                    ? 'Signing in...'
                    : 'Login'}
                </button>
              </form>

              <button
                className="forgotLink"
                onClick={() => {
                  setForgotMode(true)
                  setForgotMessage('')
                }}
              >
                Forgot password?
              </button>

              {loginError && (
                <div className="loginError">
                  {loginError}
                </div>
              )}
            </>
          ) : (
            <>
              <h1>Reset Password</h1>

              <p>
                Enter your registered email address.
              </p>

              <form onSubmit={resetPassword}>
                <input
                  type="email"
                  placeholder="Email address"
                  value={forgotEmail}
                  onChange={event =>
                    setForgotEmail(
                      event.target.value
                    )
                  }
                  required
                />

                <button className="primaryLogin">
                  Send Reset Link
                </button>
              </form>

              <button
                className="forgotLink"
                onClick={() => {
                  setForgotMode(false)
                  setForgotMessage('')
                }}
              >
                Back to Login
              </button>

              {forgotMessage && (
                <div className="loginInfo">
                  {forgotMessage}
                </div>
              )}
            </>
          )}

        </section>
      </main>
    )
  }

  const maxRMValue =
    Math.max(
      ...rmData.map(item => item.value),
      1
    )

  return (
    <main className="appShell">

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
              page === 'dashboard'
                ? 'navItem active'
                : 'navItem'
            }
            onClick={() =>
              setPage('dashboard')
            }
          >
            <LayoutDashboard size={18} />
            Dashboard
          </button>

          <button
            className={
              page === 'transactions'
                ? 'navItem active'
                : 'navItem'
            }
            onClick={() =>
              setPage('transactions')
            }
          >
            <TableProperties size={18} />
            Transaction Data
          </button>

          {isAdmin && (
            <>
              <button
                className={
                  page === 'rms'
                    ? 'navItem active'
                    : 'navItem'
                }
                onClick={() =>
                  setPage('rms')
                }
              >
                <Users size={18} />
                Manage RMs
              </button>

              <button
                className={
                  page === 'settings'
                    ? 'navItem active'
                    : 'navItem'
                }
                onClick={() =>
                  setPage('settings')
                }
              >
                <Settings size={18} />
                Settings
              </button>
            </>
          )}

        </nav>

        <div className="sidebarBottom">
          <button
            className="logoutButton"
            onClick={logout}
          >
            <LogOut size={18} />
            Logout
          </button>
        </div>

      </aside>

      <section className="mainContent">

        <header className="topHeader">

          <div>
            <h1>
              {page === 'dashboard'
                ? 'Snowball Redemption Tracker'
                : page === 'transactions'
                  ? 'Transaction Data'
                  : page === 'rms'
                    ? 'Manage Relationship Managers'
                    : 'Settings'}
            </h1>

            <p>
              {page === 'dashboard'
                ? 'Analyse transactions and monitor redemption activity'
                : page === 'transactions'
                  ? 'Upload, review and export transaction details'
                  : page === 'rms'
                    ? 'Add new RMs or mark existing RMs as inactive'
                    : 'Manage administrator access'}
            </p>
          </div>

          <button
            className="headerLogout"
            onClick={logout}
          >
            <LogOut size={16} />
            Logout
          </button>

        </header>

        {error && (
          <div className="errorBox">
            {error}

            <button
              onClick={() => setError('')}
            >
              <X size={16} />
            </button>
          </div>
        )}

        {(page === 'dashboard' ||
          page === 'transactions') && (

          <section className="toolbar">

            <div className="periodButtons">
              {['WTD', 'MTD', 'QTD', 'YTD'].map(item => (
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
                    clearDates()
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
              <option value="All">
                All RMs
              </option>

              {activeRMs.map(item => (
                <option
                  key={item}
                  value={item}
                >
                  {item}
                </option>
              ))}
            </select>

            <input
              type="date"
              value={from}
              onChange={event =>
                setFrom(event.target.value)
              }
            />

            <input
              type="date"
              value={to}
              onChange={event =>
                setTo(event.target.value)
              }
            />

            <button
              className="iconButton"
              title="Refresh"
              onClick={initialise}
            >
              <RefreshCw size={17} />
            </button>

            {page === 'transactions' && (
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
            )}

          </section>
        )}

        {message && (
          <div className="messageBox">
            {message}
          </div>
        )}

        {page === 'dashboard' && (
          <>

            <section className="metricGrid">

              <article className="metricCard redemptionCard">
                <span>Redemption</span>
                <strong>
                  {money(totals.Redemption)}
                </strong>
              </article>

              <article className="metricCard swpCard">
                <span>SWP</span>
                <strong>
                  {money(totals.SWP)}
                </strong>
              </article>

              <article className="metricCard switchCard">
                <span>Switch</span>
                <strong>
                  {money(totals.Switch)}
                </strong>
              </article>

              <article className="metricCard stpCard">
                <span>STP</span>
                <strong>
                  {money(totals.STP)}
                </strong>
              </article>

              <article className="metricCard">
                <span>Investors</span>
                <strong>
                  {totals.Investors}
                </strong>
              </article>

              <article className="metricCard">
                <span>Transactions</span>
                <strong>
                  {totals.Transactions}
                </strong>
              </article>

            </section>

            <section className="dashboardGrid">

              <article className="panel classificationPanel">
                <h2>
                  Amount by Classification
                </h2>

                <div className="donutLayout">

                  <div
                    className="donut"
                    style={{
                      background: donutBackground
                    }}
                  >
                    <div className="donutCenter">
                      <strong>
                        {totals.Transactions}
                      </strong>
                      <span>Transactions</span>
                    </div>
                  </div>

                  <div className="legend">
                    {classificationData.map(item => (
                      <div
                        className="legendRow"
                        key={item.name}
                      >
                        <span
                          className="legendDot"
                          style={{
                            backgroundColor:
                              item.color
                          }}
                        />

                        <span>
                          {item.name}
                        </span>

                        <strong>
                          {shortMoney(item.value)}
                        </strong>
                      </div>
                    ))}
                  </div>

                </div>
              </article>

              <article className="panel trendPanel">
                <h2>
                  Monthly Trend (Amount in ₹)
                </h2>

                <div className="trendLegend">
                  <span className="trendRedemption">
                    Redemption
                  </span>
                  <span className="trendSwp">
                    SWP
                  </span>
                  <span className="trendSwitch">
                    Switch
                  </span>
                  <span className="trendStp">
                    STP
                  </span>
                </div>

                <div className="trendChart">

                  {monthlyData.length === 0 ? (
                    <div className="emptyChart">
                      No transaction data
                    </div>
                  ) : (
                    monthlyData.map(item => {
                      const total =
                        item.Redemption +
                        item.SWP +
                        item.Switch +
                        item.STP

                      const maxMonth =
                        Math.max(
                          ...monthlyData.map(month =>
                            month.Redemption +
                            month.SWP +
                            month.Switch +
                            month.STP
                          ),
                          1
                        )

                      const height =
                        Math.max(
                          6,
                          (total / maxMonth) * 150
                        )

                      return (
                        <div
                          className="monthBar"
                          key={item.month}
                        >
                          <div
                            className="stack"
                            style={{
                              height:
                                `${height}px`
                            }}
                          >
                            <div
                              className="stackRedemption"
                              style={{
                                height:
                                  `${total
                                    ? (item.Redemption / total) * 100
                                    : 0}%`
                              }}
                            />

                            <div
                              className="stackSwp"
                              style={{
                                height:
                                  `${total
                                    ? (item.SWP / total) * 100
                                    : 0}%`
                              }}
                            />

                            <div
                              className="stackSwitch"
                              style={{
                                height:
                                  `${total
                                    ? (item.Switch / total) * 100
                                    : 0}%`
                              }}
                            />

                            <div
                              className="stackStp"
                              style={{
                                height:
                                  `${total
                                    ? (item.STP / total) * 100
                                    : 0}%`
                              }}
                            />
                          </div>

                          <span>
                            {item.month.slice(5)}
                          </span>
                        </div>
                      )
                    })
                  )}

                </div>
              </article>

              <article className="panel rmPanel">
                <h2>
                  Transactions by RM
                </h2>

                <div className="rmBars">

                  {rmData.length === 0 && (
                    <div className="emptyChart">
                      No RM data
                    </div>
                  )}

                  {rmData.map(item => (
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
                            width:
                              `${(item.value / maxRMValue) * 100}%`
                          }}
                        />
                      </div>

                      <strong>
                        {shortMoney(item.value)}
                      </strong>
                    </div>
                  ))}

                </div>
              </article>

              <article className="panel analysisPanel">
                <h2>
                  Classification Analysis
                </h2>

                <div className="analysisRows">

                  {classificationData.map(item => {
                    const percentage =
                      totalClassification
                        ? (
                          item.value /
                          totalClassification
                        ) * 100
                        : 0

                    return (
                      <div
                        className="analysisRow"
                        key={item.name}
                      >
                        <div>
                          <span
                            className="analysisDot"
                            style={{
                              backgroundColor:
                                item.color
                            }}
                          />

                          {item.name}
                        </div>

                        <div className="analysisTrack">
                          <div
                            style={{
                              width:
                                `${percentage}%`,
                              backgroundColor:
                                item.color
                            }}
                          />
                        </div>

                        <strong>
                          {shortMoney(item.value)}
                        </strong>
                      </div>
                    )
                  })}

                </div>
              </article>

            </section>

          </>
        )}

        {page === 'transactions' && (
          <section className="dataPanel">

            <div className="sectionTitle">
              <div>
                <h2>Transaction Details</h2>
                <p>
                  {filtered.length} transactions found
                </p>
              </div>
            </div>

            <div className="tableWrap">

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
                    .slice(0, 1000)
                    .map(item => (
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
                          {item.original_transaction_type}
                        </td>

                        <td>
                          <span
                            className={
                              `classificationBadge ${String(
                                item.classified_transaction_type
                              ).toLowerCase()}`
                            }
                          >
                            {item.classified_transaction_type}
                          </span>
                        </td>
                      </tr>
                    ))}

                </tbody>
              </table>

            </div>
          </section>
        )}

        {page === 'rms' && isAdmin && (
          <section className="managementPage">

            <article className="managementCard">

              <h2>Add New RM</h2>

              <form
                className="inlineForm"
                onSubmit={addRM}
              >
                <input
                  placeholder="Enter RM name"
                  value={newRM}
                  onChange={event =>
                    setNewRM(event.target.value)
                  }
                />

                <button>
                  <Plus size={17} />
                  Add RM
                </button>
              </form>

            </article>

            <article className="managementCard">

              <h2>
                Relationship Managers
              </h2>

              <div className="managementTable">

                {rms.map(item => (
                  <div
                    className="managementRow"
                    key={item.id}
                  >
                    <div>
                      <strong>
                        {item.name}
                      </strong>

                      <span
                        className={
                          item.is_active
                            ? 'status activeStatus'
                            : 'status inactiveStatus'
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
                        toggleRM(item)
                      }
                    >
                      {item.is_active ? (
                        <>
                          <UserX size={16} />
                          Mark Inactive
                        </>
                      ) : (
                        <>
                          <UserCheck size={16} />
                          Activate
                        </>
                      )}
                    </button>
                  </div>
                ))}

              </div>

            </article>

          </section>
        )}

        {page === 'settings' && isAdmin && (
          <section className="managementPage">

            <article className="managementCard">

              <h2>
                <KeyRound size={20} />
                Administrator Access
              </h2>

              <p className="mutedText">
                Add or deactivate users who should have
                administrative access to Manage RMs and Settings.
              </p>

              <form
                className="inlineForm"
                onSubmit={addAdmin}
              >
                <input
                  type="email"
                  placeholder="Administrator email address"
                  value={newAdmin}
                  onChange={event =>
                    setNewAdmin(event.target.value)
                  }
                />

                <button>
                  <Mail size={17} />
                  Add Admin
                </button>
              </form>

            </article>

            <article className="managementCard">

              <h2>
                Current Administrators
              </h2>

              <div className="managementTable">

                {admins.map(item => (
                  <div
                    className="managementRow"
                    key={item.id}
                  >
                    <div>
                      <strong>
                        {item.email}
                      </strong>

                      <span
                        className={
                          item.is_active
                            ? 'status activeStatus'
                            : 'status inactiveStatus'
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
                        toggleAdmin(item)
                      }
                    >
                      {item.is_active
                        ? 'Deactivate'
                        : 'Activate'}
                    </button>
                  </div>
                ))}

              </div>

            </article>

          </section>
        )}

      </section>
    </main>
  )
}

createRoot(
  document.getElementById('root')
).render(<App />)
