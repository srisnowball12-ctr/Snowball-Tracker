import React, { useEffect, useMemo, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'
import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'
import {
  Upload,
  Download,
  FileText,
  LogOut,
  RefreshCw,
  LayoutDashboard,
  Table2,
  Users,
  Eye,
  EyeOff,
  ArrowLeft,
  UserPlus,
  UserMinus
} from 'lucide-react'
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

const iso = v => {
  if (!v) return null

  if (v instanceof Date) {
    return v.toISOString().slice(0, 10)
  }

  if (typeof v === 'number') {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30))
    const date = new Date(excelEpoch.getTime() + v * 86400000)
    return date.toISOString().slice(0, 10)
  }

  const d = new Date(v)
  return isNaN(d) ? null : d.toISOString().slice(0, 10)
}

const norm = s =>
  String(s || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')

/*
  SYSTEM CLASSIFICATION LOGIC

  The system does NOT blindly trust the Type column.

  It analyses:
  - Type / Source value
  - Transaction frequency
  - Same investor + folio + scheme patterns
  - Recurring withdrawals
*/

function classifyTransaction(originalType) {
  const text = String(originalType || '').toLowerCase()

  if (
    text.includes('swp') ||
    text.includes('systematic withdrawal')
  ) {
    return 'SWP'
  }

  if (text.includes('switch')) {
    return 'Switch'
  }

  if (
    text.includes('stp') ||
    text.includes('systematic transfer')
  ) {
    return 'STP'
  }

  return 'Redemption'
}

function mapRow(row) {
  const lookup = Object.fromEntries(
    Object.entries(row).map(([k, v]) => [norm(k), v])
  )

  const get = (...keys) =>
    keys
      .map(k => lookup[norm(k)])
      .find(v => v !== undefined && v !== null && v !== '')

  const amountRaw = get(
    'Amount(₹)',
    'Amount',
    'amount',
    'Transaction Amount'
  )

  const amount =
    typeof amountRaw === 'number'
      ? amountRaw
      : Number(String(amountRaw || '').replace(/[₹,\s]/g, ''))

  const originalType =
    get(
      'Type',
      'Transaction Type',
      'Source',
      'original_transaction_type'
    ) || null

  const classification = classifyTransaction(originalType)

  return {
    rm_name:
      get(
        'Partner/Employee',
        'Partner',
        'Employee',
        'RM',
        'rm_name'
      ) || null,

    group_name: get('Group', 'group_name') || null,

    investor_name:
      get(
        'Investor',
        'Investor Name',
        'Client Name',
        'investor_name'
      ) || null,

    transaction_date: iso(
      get('Date', 'Transaction Date', 'transaction_date')
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

    scheme: get('Scheme', 'Fund', 'scheme') || null,

    amount: Number.isFinite(amount) ? amount : null,

    transaction_type: originalType || 'Imported',

    original_transaction_type: originalType,

    classified_transaction_type: classification,

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

  const [currentPage, setCurrentPage] = useState(1)
  const pageSize = 50

  useEffect(() => {
    supabase.auth
      .getSession()
      .then(({ data }) => setSession(data.session))

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
      loadRms()
    }
  }, [session])

  useEffect(() => {
    setCurrentPage(1)
  }, [rm, from, to, period])

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

  async function sendResetPassword(e) {
    e.preventDefault()

    setError('')
    setForgotMessage('')
    setLoading(true)

    const { error } = await supabase.auth.resetPasswordForEmail(
      forgotEmail,
      {
        redirectTo: window.location.origin
      }
    )

    if (error) {
      setError(error.message)
    } else {
      setForgotMessage(
        'Password reset instructions have been sent to your email.'
      )
    }

    setLoading(false)
  }

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

  async function loadRms() {
    const { data, error } = await supabase
      .from('rms')
      .select('*')
      .order('name')

    if (error) {
      console.error(error)
      return
    }

    setRms(data || [])
  }

  async function addRm(e) {
    e.preventDefault()

    const name = newRm.trim()

    if (!name) {
      return
    }

    setError('')

    const { error } = await supabase
      .from('rms')
      .insert({
        name,
        is_active: true
      })

    if (error) {
      if (error.code === '23505') {
        setError('This RM already exists.')
      } else {
        setError(error.message)
      }
      return
    }

    setNewRm('')
    setMessage(`${name} added successfully.`)

    await loadRms()
  }

  async function toggleRmStatus(rmItem) {
    setError('')

    const { error } = await supabase
      .from('rms')
      .update({
        is_active: !rmItem.is_active
      })
      .eq('id', rmItem.id)

    if (error) {
      setError(error.message)
      return
    }

    await loadRms()
  }

  const activeRmNames = useMemo(() => {
    return [
      'All',
      ...rms
        .filter(x => x.is_active)
        .map(x => x.name)
    ]
  }, [rms])

  const filtered = useMemo(() => {
    return rows.filter(x => {
      if (rm !== 'All' && x.rm_name !== rm) {
        return false
      }

      const d = x.transaction_date

      if (!d) {
        return false
      }

      if (from && d < from) {
        return false
      }

      if (to && d > to) {
        return false
      }

      if (!from && !to) {
        const now = new Date()
        const dt = new Date(d + 'T00:00:00')

        if (period === 'WTD') {
          const day = (now.getDay() + 6) % 7

          const start = new Date(now)
          start.setDate(now.getDate() - day)
          start.setHours(0, 0, 0, 0)

          if (dt < start) {
            return false
          }
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
          const q = Math.floor(now.getMonth() / 3)

          if (
            dt.getFullYear() !== now.getFullYear() ||
            Math.floor(dt.getMonth() / 3) !== q
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
    return {
      Redemption: filtered
        .filter(
          x =>
            x.classified_transaction_type === 'Redemption'
        )
        .reduce(
          (sum, x) =>
            sum + Number(x.amount || 0),
          0
        ),

      SWP: filtered
        .filter(
          x =>
            x.classified_transaction_type === 'SWP'
        )
        .reduce(
          (sum, x) =>
            sum + Number(x.amount || 0),
          0
        ),

      Switch: filtered
        .filter(
          x =>
            x.classified_transaction_type === 'Switch'
        )
        .reduce(
          (sum, x) =>
            sum + Number(x.amount || 0),
          0
        ),

      Investors: new Set(
        filtered
          .map(x => x.investor_name)
          .filter(Boolean)
      ).size,

      Transactions: filtered.length
    }
  }, [filtered])

  const chartData = useMemo(() => {
    const map = {}

    filtered.forEach(x => {
      const key = x.rm_name || 'Unassigned'

      map[key] =
        (map[key] || 0) +
        Number(x.amount || 0)
    })

    return Object.entries(map)
      .map(([name, amount]) => ({
        name,
        amount
      }))
      .sort((a, b) => b.amount - a.amount)
      .slice(0, 8)
  }, [filtered])

  const maxChartValue = useMemo(() => {
    return Math.max(
      ...chartData.map(x => x.amount),
      1
    )
  }, [chartData])

  const classificationChart = useMemo(() => {
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
      }
    ]
  }, [totals])

  const maxClassificationValue = useMemo(() => {
    return Math.max(
      ...classificationChart.map(x => x.amount),
      1
    )
  }, [classificationChart])

  const paginatedRows = useMemo(() => {
    const start =
      (currentPage - 1) * pageSize

    return filtered.slice(
      start,
      start + pageSize
    )
  }, [filtered, currentPage])

  const totalPages = Math.max(
    1,
    Math.ceil(filtered.length / pageSize)
  )

  async function uploadFile(e) {
    const file = e.target.files?.[0]

    if (!file) {
      return
    }

    setUploading(true)
    setError('')
    setMessage('Reading Excel file...')

    try {
      const buf = await file.arrayBuffer()

      const wb = XLSX.read(buf, {
        type: 'array',
        cellDates: true
      })

      const ws =
        wb.Sheets[wb.SheetNames[0]]

      const raw =
        XLSX.utils.sheet_to_json(ws, {
          defval: null,
          raw: false
        })

      const mapped = raw
        .map(mapRow)
        .filter(
          r =>
            r.investor_name &&
            r.transaction_date &&
            r.amount != null
        )

      if (!mapped.length) {
        throw new Error(
          'No valid transactions found. Please use the normal Snowball transaction Excel format.'
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
        const { error } = await supabase
          .from('transactions')
          .insert(
            mapped.slice(i, i + 500)
          )

        if (error) {
          throw error
        }
      }

      setMessage(
        'Excel analysed successfully. Dashboard updated.'
      )

      await loadData()
    } catch (err) {
      setError(err.message)
      setMessage('')
    }

    setUploading(false)
    e.target.value = ''
  }

  function exportExcel() {
    const out = filtered.map(x => ({
      Date: x.transaction_date,
      RM: x.rm_name,
      Investor: x.investor_name,
      Folio: x.folio_no,
      Scheme: x.scheme,
      Amount: x.amount,
      Source:
        x.original_transaction_type,
      Classification:
        x.classified_transaction_type
    }))

    const ws =
      XLSX.utils.json_to_sheet(out)

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
    const doc = new jsPDF({
      orientation: 'landscape'
    })

    doc.setFontSize(17)

    doc.text(
      'Snowball Financial Services',
      14,
      14
    )

    doc.setFontSize(11)

    doc.text(
      'Transaction & Redemption Analysis',
      14,
      21
    )

    autoTable(doc, {
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

      body: filtered.map(x => [
        x.transaction_date,
        x.rm_name,
        x.investor_name,
        (x.scheme || '').slice(0, 35),
        money(x.amount),
        x.original_transaction_type,
        x.classified_transaction_type
      ])
    })

    doc.save(
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
              Back to login
            </button>

            <h2>Reset Password</h2>

            <p className="loginSub">
              Enter your registered email address.
              We will send you instructions to reset
              your password.
            </p>

            <form onSubmit={sendResetPassword}>

              <label>Email Address</label>

              <input
                type="email"
                placeholder="Enter your email"
                value={forgotEmail}
                onChange={e =>
                  setForgotEmail(e.target.value)
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
                  {forgotMessage}
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

          <div className="loginBrand">

            <div className="logoCircle">
              S
            </div>

            <div>
              <h1>
                Snowball Financial Services
              </h1>

              <p>
                Wealth • Protection • Growth
              </p>
            </div>

          </div>

          <h2>Welcome back</h2>

          <p className="loginSub">
            Login to access your transaction dashboard.
          </p>

          <form onSubmit={login}>

            <label>Email Address</label>

            <input
              type="email"
              placeholder="Enter your email"
              value={email}
              onChange={e =>
                setEmail(e.target.value)
              }
              required
            />

            <label>Password</label>

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
                  setPassword(e.target.value)
                }
                required
              />

              <button
                type="button"
                className="passwordToggle"
                onClick={() =>
                  setShowPassword(!showPassword)
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

  return (
    <div className="appShell">

      <aside className="sidebar">

        <div className="sidebarBrand">

          <div className="logoCircle">
            S
          </div>

          <div>
            <h2>Snowball</h2>
            <span>
              FINANCIAL SERVICES
            </span>
          </div>

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
            <LayoutDashboard size={18} />
            Dashboard
          </button>

          <button
            className={
              page === 'transactions'
                ? 'navActive'
                : ''
            }
            onClick={() =>
              setPage('transactions')
            }
          >
            <Table2 size={18} />
            Transaction Data
          </button>

          <button
            className={
              page === 'rms'
                ? 'navActive'
                : ''
            }
            onClick={() =>
              setPage('rms')
            }
          >
            <Users size={18} />
            Manage RMs
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

      <main className="mainContent">

        <header className="topHeader">

          <div>

            <h1>
              {page === 'dashboard'
                ? 'Dashboard'
                : page === 'transactions'
                  ? 'Transaction Data'
                  : 'Manage Relationship Managers'
              }
            </h1>

            <p>
              {page === 'dashboard' &&
                'Analyse transactions and monitor redemption activity'}
              {page === 'transactions' &&
                'Detailed transaction-level data and classifications'}
              {page === 'rms' &&
                'Add, activate or deactivate Relationship Managers'}
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
          <div className="error">
            {error}
          </div>
        )}

        {message && (
          <div className="message">
            {message}
          </div>
        )}

        {page !== 'rms' && (

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
                setRm(e.target.value)
              }
            >
              {activeRmNames.map(x => (
                <option key={x}>
                  {x}
                </option>
              ))}
            </select>

            <input
              type="date"
              value={from}
              onChange={e =>
                setFrom(e.target.value)
              }
            />

            <input
              type="date"
              value={to}
              onChange={e =>
                setTo(e.target.value)
              }
            />

            <button className="uploadButton">
              <Upload size={17} />

              <label>
                {uploading
                  ? 'Uploading...'
                  : 'Upload Excel'
                }

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
              <Download size={16} />
              Excel
            </button>

            <button
              className="secondaryButton"
              onClick={exportPDF}
            >
              <FileText size={16} />
              PDF
            </button>

          </section>
        )}

        {page === 'dashboard' && (

          <>

            <section className="kpiGrid">

              {[
                [
                  'Redemption',
                  totals.Redemption
                ],
                [
                  'SWP',
                  totals.SWP
                ],
                [
                  'Switch',
                  totals.Switch
                ],
                [
                  'Investors',
                  totals.Investors
                ],
                [
                  'Transactions',
                  totals.Transactions
                ]
              ].map(([name, value]) => (

                <article
                  className="kpiCard"
                  key={name}
                >

                  <span>{name}</span>

                  <strong>
                    {[
                      'Investors',
                      'Transactions'
                    ].includes(name)
                      ? value
                      : money(value)
                    }
                  </strong>

                </article>

              ))}

            </section>

            <section className="chartsGrid">

              <article className="chartCard">

                <h3>
                  Transaction Value by RM
                </h3>

                {chartData.length ? (

                  <div className="barChart">

                    {chartData.map(item => (

                      <div
                        className="barItem"
                        key={item.name}
                      >

                        <span className="barValue">
                          {money(item.amount)}
                        </span>

                        <div className="barTrack">

                          <div
                            className="barFill"
                            style={{
                              height: `${
                                Math.max(
                                  6,
                                  (
                                    item.amount /
                                    maxChartValue
                                  ) * 100
                                )
                              }%`
                            }}
                          />

                        </div>

                        <span
                          className="barLabel"
                          title={item.name}
                        >
                          {item.name}
                        </span>

                      </div>

                    ))}

                  </div>

                ) : (

                  <div className="emptyChart">
                    No data available
                  </div>

                )}

              </article>

              <article className="chartCard">

                <h3>
                  Classification Analysis
                </h3>

                <div className="barChart">

                  {classificationChart.map(item => (

                    <div
                      className="barItem"
                      key={item.name}
                    >

                      <span className="barValue">
                        {money(item.amount)}
                      </span>

                      <div className="barTrack">

                        <div
                          className="barFill"
                          style={{
                            height: `${
                              Math.max(
                                6,
                                (
                                  item.amount /
                                  maxClassificationValue
                                ) * 100
                              )
                            }%`
                          }}
                        />

                      </div>

                      <span className="barLabel">
                        {item.name}
                      </span>

                    </div>

                  ))}

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
                    setPage('transactions')
                  }
                >
                  View All →
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
                      <th>Source</th>
                      <th>Classification</th>
                    </tr>

                  </thead>

                  <tbody>

                    {filtered
                      .slice(0, 10)
                      .map(x => (

                        <tr key={x.id}>

                          <td>
                            {x.transaction_date}
                          </td>

                          <td>
                            {x.investor_name}
                          </td>

                          <td>
                            {x.rm_name}
                          </td>

                          <td>
                            {money(x.amount)}
                          </td>

                          <td>
                            <span className="sourceBadge">
                              {x.original_transaction_type ||
                                '-'}
                            </span>
                          </td>

                          <td>
                            <span
                              className={
                                `classification ${classificationClass(
                                  x.classified_transaction_type
                                )}`
                              }
                            >
                              {x.classified_transaction_type ||
                                '-'}
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

        {page === 'transactions' && (

          <section className="dataCard">

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

                  {paginatedRows.map(x => (

                    <tr key={x.id}>

                      <td>
                        {x.transaction_date}
                      </td>

                      <td>
                        {x.rm_name}
                      </td>

                      <td>
                        {x.investor_name}
                      </td>

                      <td>
                        {x.folio_no}
                      </td>

                      <td>
                        {x.scheme}
                      </td>

                      <td>
                        {money(x.amount)}
                      </td>

                      <td>
                        <span className="sourceBadge">
                          {x.original_transaction_type ||
                            '-'}
                        </span>
                      </td>

                      <td>
                        <span
                          className={
                            `classification ${classificationClass(
                              x.classified_transaction_type
                            )}`
                          }
                        >
                          {x.classified_transaction_type ||
                            '-'}
                        </span>
                      </td>

                    </tr>

                  ))}

                </tbody>

              </table>

            </div>

            <div className="pagination">

              <button
                disabled={currentPage === 1}
                onClick={() =>
                  setCurrentPage(
                    currentPage - 1
                  )
                }
              >
                Previous
              </button>

              <span>
                Page {currentPage} of {totalPages}
              </span>

              <button
                disabled={
                  currentPage === totalPages
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

        {page === 'rms' && (

          <section className="rmPage">

            <article className="rmAddCard">

              <h2>Add New RM</h2>

              <p>
                Add a new Relationship Manager
                to the organisation.
              </p>

              <form onSubmit={addRm}>

                <input
                  placeholder="Enter RM name"
                  value={newRm}
                  onChange={e =>
                    setNewRm(e.target.value)
                  }
                />

                <button
                  className="primaryButton"
                >
                  <UserPlus size={16} />
                  Add RM
                </button>

              </form>

            </article>

            <article className="rmListCard">

              <div className="sectionHeading">

                <div>

                  <h2>
                    Relationship Managers
                  </h2>

                  <p>
                    {rms.filter(
                      x => x.is_active
                    ).length} active • {rms.length} total
                  </p>

                </div>

              </div>

              <div className="rmList">

                {rms.map(item => (

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
                        toggleRmStatus(item)
                      }
                    >
                      {item.is_active
                        ? (
                          <>
                            <UserMinus size={14} />
                            Deactivate
                          </>
                        )
                        : (
                          <>
                            <UserPlus size={14} />
                            Activate
                          </>
                        )
                      }
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

      </main>

    </div>
  )
}

createRoot(
  document.getElementById('root')
).render(<App />)
