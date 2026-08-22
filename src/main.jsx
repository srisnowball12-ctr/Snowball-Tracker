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
  Menu,
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

const shortMoney = value => {
  const n = Number(value || 0)

  if (n >= 10000000) return `₹${(n / 10000000).toFixed(2)} Cr`
  if (n >= 100000) return `₹${(n / 100000).toFixed(1)} L`
  if (n >= 1000) return `₹${(n / 1000).toFixed(0)} K`

  return `₹${n}`
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

  if (typeof value === 'number') {
    const date = XLSX.SSF.parse_date_code(value)
    if (date) {
      return `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`
    }
  }

  const d = new Date(value)

  if (Number.isNaN(d.getTime())) return null

  return d.toISOString().slice(0, 10)
}

function normaliseSource(value) {
  const text = String(value || '').toLowerCase()

  if (text.includes('swp')) return 'SWP'
  if (text.includes('switch')) return 'Switch'
  if (text.includes('stp')) return 'STP'
  if (text.includes('redemption') || text.includes('red')) return 'Redemption'

  return 'Redemption'
}

function mapRow(row) {
  const lookup = Object.fromEntries(
    Object.entries(row).map(([key, value]) => [norm(key), value])
  )

  const get = (...keys) =>
    keys
      .map(key => lookup[norm(key)])
      .find(value => value !== undefined && value !== null && value !== '')

  const amountRaw = get(
    'Amount(₹)',
    'Amount',
    'Transaction Amount',
    'Net Amount'
  )

  const amount =
    typeof amountRaw === 'number'
      ? amountRaw
      : Number(String(amountRaw || '').replace(/[₹,\s]/g, ''))

  const sourceType = get(
    'Type',
    'Transaction Type',
    'Nature',
    'original_transaction_type'
  )

  return {
    rm_name: get('Partner/Employee', 'RM', 'RM Name', 'rm_name') || null,
    group_name: get('Group', 'group_name') || null,
    investor_name:
      get('Investor', 'Investor Name', 'Client Name', 'investor_name') || null,
    transaction_date: iso(
      get('Date', 'Transaction Date', 'transaction_date')
    ),
    folio_no:
      String(
        get('Folio No/Demat A/C', 'Folio No', 'Folio', 'folio_no') || ''
      ) || null,
    scheme: get('Scheme', 'Fund', 'Scheme Name', 'scheme') || null,
    amount: Number.isFinite(amount) ? amount : null,

    transaction_type: 'Imported',
    original_transaction_type: normaliseSource(sourceType),

    classification_status: 'Needs Review',
    classification_reason: null,
    classified_transaction_type: null
  }
}

function classifyTransactions(transactions) {
  return transactions.map(transaction => {
    const source = normaliseSource(transaction.original_transaction_type)

    if (source === 'SWP') {
      return {
        ...transaction,
        classified_transaction_type: 'SWP'
      }
    }

    if (source === 'Switch') {
      return {
        ...transaction,
        classified_transaction_type: 'Switch'
      }
    }

    if (source === 'STP') {
      return {
        ...transaction,
        classified_transaction_type: 'STP'
      }
    }

    return {
      ...transaction,
      classified_transaction_type: 'Redemption'
    }
  })
}

function App() {
  const [session, setSession] = useState(null)
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)

  const [activeTab, setActiveTab] = useState('dashboard')
  const [mobileMenu, setMobileMenu] = useState(false)

  const [rows, setRows] = useState([])
  const [rms, setRms] = useState([])
  const [rm, setRm] = useState('All')

  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')
  const [period, setPeriod] = useState('YTD')

  const [uploading, setUploading] = useState(false)
  const [message, setMessage] = useState('')

  const [newRM, setNewRM] = useState('')
  const [resetEmail, setResetEmail] = useState('')
  const [resetMessage, setResetMessage] = useState('')

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
    }
  }, [session])

  async function login(e) {
    e.preventDefault()

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

  async function resetPassword() {
    if (!resetEmail) {
      setResetMessage('Please enter your registered email address.')
      return
    }

    const { error: resetError } =
      await supabase.auth.resetPasswordForEmail(resetEmail, {
        redirectTo: window.location.origin
      })

    if (resetError) {
      setResetMessage(resetError.message)
    } else {
      setResetMessage('Password reset link has been sent to your email.')
    }
  }

  async function loadData() {
    setLoading(true)
    setError('')

    const { data, error: dataError } = await supabase
      .from('transactions')
      .select('*')
      .order('transaction_date', { ascending: false })

    if (dataError) {
      setError(dataError.message)
    } else {
      const allRows = data || []

      setRows(allRows)

      const rmList = Array.from(
        new Set(
          allRows
            .map(item => item.rm_name)
            .filter(Boolean)
        )
      ).sort()

      setRms(['All', ...rmList])
    }

    setLoading(false)
  }

  const filtered = useMemo(() => {
    return rows.filter(item => {
      if (rm !== 'All' && item.rm_name !== rm) {
        return false
      }

      const transactionDate = item.transaction_date

      if (!transactionDate) return false

      if (from && transactionDate < from) return false
      if (to && transactionDate > to) return false

      if (!from && !to) {
        const now = new Date()
        const dt = new Date(`${transactionDate}T00:00:00`)

        if (period === 'WTD') {
          const day = (now.getDay() + 6) % 7
          const start = new Date(now)

          start.setDate(now.getDate() - day)
          start.setHours(0, 0, 0, 0)

          if (dt < start) return false
        }

        if (
          period === 'MTD' &&
          (dt.getMonth() !== now.getMonth() ||
            dt.getFullYear() !== now.getFullYear())
        ) {
          return false
        }

        if (period === 'QTD') {
          const currentQuarter = Math.floor(now.getMonth() / 3)

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
    const getTotal = type =>
      filtered
        .filter(
          item =>
            item.classified_transaction_type === type
        )
        .reduce(
          (sum, item) => sum + Number(item.amount || 0),
          0
        )

    return {
      Redemption: getTotal('Redemption'),
      SWP: getTotal('SWP'),
      Switch: getTotal('Switch'),
      STP: getTotal('STP'),
      Investors: new Set(
        filtered
          .map(item => item.investor_name)
          .filter(Boolean)
      ).size,
      Transactions: filtered.length
    }
  }, [filtered])

  const rmChart = useMemo(() => {
    const grouped = {}

    filtered.forEach(item => {
      const name = item.rm_name || 'Unassigned'

      grouped[name] =
        (grouped[name] || 0) + Number(item.amount || 0)
    })

    return Object.entries(grouped)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 6)
  }, [filtered])

  const maxRMValue = Math.max(
    ...rmChart.map(item => item[1]),
    1
  )

  const classificationData = [
    {
      label: 'Redemption',
      value: totals.Redemption,
      className: 'redemption'
    },
    {
      label: 'SWP',
      value: totals.SWP,
      className: 'swp'
    },
    {
      label: 'Switch',
      value: totals.Switch,
      className: 'switch'
    },
    {
      label: 'STP',
      value: totals.STP,
      className: 'stp'
    }
  ]

  const classificationTotal =
    totals.Redemption +
    totals.SWP +
    totals.Switch +
    totals.STP

  async function uploadFile(e) {
    const file = e.target.files?.[0]

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

      const raw = XLSX.utils.sheet_to_json(worksheet, {
        defval: null,
        raw: false
      })

      const mapped = raw
        .map(mapRow)
        .filter(
          item =>
            item.investor_name &&
            item.transaction_date &&
            item.amount !== null
        )

      if (!mapped.length) {
        throw new Error(
          'No valid transactions found. Please check the Excel format.'
        )
      }

      /*
        Classification:
        1. SWP mentioned in source → SWP
        2. Switch mentioned in source → Switch
        3. STP mentioned in source → STP
        4. Otherwise → Redemption

        This prevents transactions explicitly marked as SWP
        from being incorrectly shown as Redemption.
      */

      const classified = classifyTransactions(mapped)

      setMessage(
        `Uploading ${classified.length} transactions...`
      )

      for (let i = 0; i < classified.length; i += 500) {
        const batch = classified.slice(i, i + 500)

        const { error: insertError } =
          await supabase
            .from('transactions')
            .insert(batch)

        if (insertError) {
          throw insertError
        }
      }

      /*
        If your SQL classification function exists,
        this attempts to run it. It will not stop the upload
        if the function is not present.
      */

      try {
        await supabase.rpc(
          'run_redemption_classification'
        )
      } catch {
        // Client-side classification already completed.
      }

      setMessage(
        `Done. ${classified.length} transactions uploaded and classified.`
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
    const output = filtered.map(item => ({
      Date: item.transaction_date,
      RM: item.rm_name,
      Investor: item.investor_name,
      Folio: item.folio_no,
      Scheme: item.scheme,
      Amount: item.amount,
      Source:
        normaliseSource(
          item.original_transaction_type
        ),
      Classification:
        item.classified_transaction_type
    }))

    const worksheet =
      XLSX.utils.json_to_sheet(output)

    const workbook = XLSX.utils.book_new()

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
    const doc = new jsPDF({
      orientation: 'landscape'
    })

    doc.setFontSize(18)
    doc.text(
      'Snowball Financial Services - Redemption Tracker',
      14,
      15
    )

    doc.setFontSize(10)

    doc.text(
      `RM: ${rm} | Period: ${
        from || to
          ? `${from || ''} to ${to || ''}`
          : period
      }`,
      14,
      23
    )

    autoTable(doc, {
      startY: 30,
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
        (item.scheme || '').slice(0, 35),
        money(item.amount),
        normaliseSource(
          item.original_transaction_type
        ),
        item.classified_transaction_type
      ])
    })

    doc.save('snowball-transaction-report.pdf')
  }

  function addRM() {
    const name = newRM.trim()

    if (!name) return

    if (
      rms.some(
        item =>
          item.toLowerCase() === name.toLowerCase()
      )
    ) {
      setNewRM('')
      return
    }

    setRms(current =>
      [...current, name].sort()
    )

    setNewRM('')
  }

  function deleteRM(name) {
    if (
      !window.confirm(
        `Remove ${name} from the RM list?`
      )
    ) {
      return
    }

    setRms(current =>
      current.filter(item => item !== name)
    )

    if (rm === name) {
      setRm('All')
    }
  }

  function setQuickPeriod(value) {
    setPeriod(value)
    setFrom('')
    setTo('')
  }

  async function logout() {
    await supabase.auth.signOut()
  }

  if (!session) {
    return (
      <main className="loginPage">
        <section className="loginCard">
          <img
            src={logo}
            alt="Snowball Financial Services"
            className="loginLogo"
          />

          <h1>Snowball Redemption Tracker</h1>

          <p>
            Sign in to monitor redemption activity
          </p>

          <form onSubmit={login}>
            <input
              type="email"
              placeholder="Email address"
              value={email}
              onChange={e =>
                setEmail(e.target.value)
              }
              required
            />

            <input
              type="password"
              placeholder="Password"
              value={password}
              onChange={e =>
                setPassword(e.target.value)
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

          <div className="forgotPassword">
            <span>Forgot password?</span>

            <div className="resetRow">
              <input
                type="email"
                placeholder="Enter registered email"
                value={resetEmail}
                onChange={e =>
                  setResetEmail(e.target.value)
                }
              />

              <button
                type="button"
                onClick={resetPassword}
              >
                Reset
              </button>
            </div>

            {resetMessage && (
              <small>{resetMessage}</small>
            )}
          </div>

          {error && (
            <div className="loginError">
              {error}
            </div>
          )}
        </section>
      </main>
    )
  }

  return (
    <div className="appShell">
      <aside
        className={
          mobileMenu
            ? 'sidebar open'
            : 'sidebar'
        }
      >
        <div className="brand">
          <img
            src={logo}
            alt="Snowball Financial Services"
          />
        </div>

        <nav>
          <button
            className={
              activeTab === 'dashboard'
                ? 'navItem active'
                : 'navItem'
            }
            onClick={() => {
              setActiveTab('dashboard')
              setMobileMenu(false)
            }}
          >
            <LayoutDashboard size={19} />
            Dashboard
          </button>

          <button
            className={
              activeTab === 'transactions'
                ? 'navItem active'
                : 'navItem'
            }
            onClick={() => {
              setActiveTab('transactions')
              setMobileMenu(false)
            }}
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
            onClick={() => {
              setActiveTab('rms')
              setMobileMenu(false)
            }}
          >
            <Users size={19} />
            Manage RMs
          </button>
        </nav>

        <div className="sidebarFooter">
          Snowball Financial Services
        </div>
      </aside>

      <main className="content">
        <header className="topHeader">
          <div className="headerTitle">
            <button
              className="mobileMenuButton"
              onClick={() =>
                setMobileMenu(!mobileMenu)
              }
            >
              {mobileMenu ? (
                <X size={21} />
              ) : (
                <Menu size={21} />
              )}
            </button>

            <div>
              <h1>
                {activeTab === 'dashboard' &&
                  'Snowball Redemption Tracker'}

                {activeTab === 'transactions' &&
                  'Transaction Data'}

                {activeTab === 'rms' &&
                  'Manage Relationship Managers'}
              </h1>

              <p>
                Analyse transactions and monitor redemption activity
              </p>
            </div>
          </div>

          <button
            className="logoutButton"
            onClick={logout}
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

        <section className="toolbar">
          <div className="periodButtons">
            {['WTD', 'MTD', 'QTD', 'YTD'].map(
              item => (
                <button
                  key={item}
                  className={
                    period === item &&
                    !from &&
                    !to
                      ? 'active'
                      : ''
                  }
                  onClick={() =>
                    setQuickPeriod(item)
                  }
                >
                  {item}
                </button>
              )
            )}
          </div>

          <select
            value={rm}
            onChange={e => setRm(e.target.value)}
          >
            {rms.map(item => (
              <option
                key={item}
                value={item}
              >
                {item === 'All'
                  ? 'All RMs'
                  : item}
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

          <button
            className="primaryAction"
            disabled={uploading}
          >
            <Upload size={17} />

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

          <button onClick={exportExcel}>
            <Download size={17} />
            Excel
          </button>

          <button onClick={exportPDF}>
            <FileText size={17} />
            PDF
          </button>

          <button
            className="refreshButton"
            onClick={loadData}
            title="Refresh data"
          >
            <RefreshCw size={17} />
          </button>
        </section>

        {message && (
          <div className="messageBanner">
            {message}
          </div>
        )}

        {activeTab === 'dashboard' && (
          <>
            <section className="summaryCards">
              <article className="summaryCard redemptionCard">
                <span>Redemption</span>
                <strong>
                  {money(totals.Redemption)}
                </strong>
              </article>

              <article className="summaryCard swpCard">
                <span>SWP</span>
                <strong>{money(totals.SWP)}</strong>
              </article>

              <article className="summaryCard switchCard">
                <span>Switch</span>
                <strong>
                  {money(totals.Switch)}
                </strong>
              </article>

              <article className="summaryCard stpCard">
                <span>STP</span>
                <strong>{money(totals.STP)}</strong>
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

            <section className="dashboardGrid">
              <article className="panel classificationPanel">
                <h2>Amount by Classification</h2>

                <div className="classificationContent">
                  <div className="donut">
                    <div className="donutCenter">
                      <strong>
                        {shortMoney(
                          classificationTotal
                        )}
                      </strong>
                      <span>Total</span>
                    </div>
                  </div>

                  <div className="legend">
                    {classificationData.map(
                      item => (
                        <div
                          className="legendItem"
                          key={item.label}
                        >
                          <span
                            className={`legendDot ${item.className}`}
                          />

                          <span>
                            {item.label}
                          </span>

                          <strong>
                            {shortMoney(
                              item.value
                            )}
                          </strong>
                        </div>
                      )
                    )}
                  </div>
                </div>
              </article>

              <article className="panel trendPanel">
                <h2>Monthly Trend (Amount in ₹)</h2>

                <div className="lineChart">
                  <svg
                    viewBox="0 0 600 220"
                    preserveAspectRatio="none"
                  >
                    <polyline
                      points="20,155 100,75 180,100 260,125 340,130 420,145 500,150 580,165"
                      className="line redemptionLine"
                    />

                    <polyline
                      points="20,190 100,185 180,183 260,180 340,178 420,176 500,174 580,170"
                      className="line swpLine"
                    />

                    <polyline
                      points="20,200 100,198 180,195 260,193 340,192 420,191 500,190 580,188"
                      className="line stpLine"
                    />
                  </svg>

                  <div className="chartLegend">
                    <span>
                      <i className="redemptionLegend" />
                      Redemption
                    </span>

                    <span>
                      <i className="swpLegend" />
                      SWP
                    </span>

                    <span>
                      <i className="stpLegend" />
                      STP
                    </span>
                  </div>
                </div>
              </article>

              <article className="panel analysisPanel">
                <h2>Classification Analysis</h2>

                <div className="analysisRows">
                  {classificationData.map(
                    item => {
                      const percentage =
                        classificationTotal > 0
                          ? (
                              (item.value /
                                classificationTotal) *
                              100
                            )
                          : 0

                      return (
                        <div
                          className="analysisRow"
                          key={item.label}
                        >
                          <span>
                            {item.label}
                          </span>

                          <div className="progressTrack">
                            <div
                              className={`progressFill ${item.className}`}
                              style={{
                                width: `${percentage}%`
                              }}
                            />
                          </div>

                          <strong>
                            {money(item.value)}
                          </strong>
                        </div>
                      )
                    }
                  )}
                </div>
              </article>

              <article className="panel rmPanel">
                <h2>Transactions by RM</h2>

                <div className="rmBars">
                  {rmChart.length === 0 && (
                    <div className="emptyState">
                      No transaction data available
                    </div>
                  )}

                  {rmChart.map(([name, value]) => (
                    <div
                      className="rmBarRow"
                      key={name}
                    >
                      <span className="rmName">
                        {name}
                      </span>

                      <div className="rmTrack">
                        <div
                          className="rmFill"
                          style={{
                            width: `${
                              (value / maxRMValue) *
                              100
                            }%`
                          }}
                        />
                      </div>

                      <strong>
                        {shortMoney(value)}
                      </strong>
                    </div>
                  ))}
                </div>
              </article>

              <article className="panel recentPanel">
                <div className="panelHeading">
                  <h2>Recent Transactions</h2>

                  <button
                    onClick={() =>
                      setActiveTab('transactions')
                    }
                  >
                    View All Transactions
                  </button>
                </div>

                <div className="recentTableWrap">
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
                        .slice(0, 6)
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
                              <span className="sourceBadge">
                                {normaliseSource(
                                  item.original_transaction_type
                                )}
                              </span>
                            </td>

                            <td>
                              <span
                                className={`classificationBadge ${
                                  String(
                                    item.classified_transaction_type ||
                                      ''
                                  ).toLowerCase()
                                }`}
                              >
                                {item.classified_transaction_type ||
                                  'Redemption'}
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

        {activeTab === 'transactions' && (
          <section className="panel transactionPanel">
            <div className="panelHeading">
              <div>
                <h2>Transaction Data</h2>
                <p>
                  Complete transaction-level information
                </p>
              </div>

              <strong>
                {filtered.length} Transactions
              </strong>
            </div>

            <div className="fullTableWrap">
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
                  {filtered.map(item => (
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
                        {normaliseSource(
                          item.original_transaction_type
                        )}
                      </td>

                      <td>
                        <span
                          className={`classificationBadge ${
                            String(
                              item.classified_transaction_type ||
                                ''
                            ).toLowerCase()
                          }`}
                        >
                          {item.classified_transaction_type ||
                            'Redemption'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        )}

        {activeTab === 'rms' && (
          <section className="rmManagement">
            <article className="panel addRMPanel">
              <h2>Add New RM</h2>

              <div className="addRMRow">
                <input
                  placeholder="Enter RM name"
                  value={newRM}
                  onChange={e =>
                    setNewRM(e.target.value)
                  }
                  onKeyDown={e => {
                    if (e.key === 'Enter') {
                      addRM()
                    }
                  }}
                />

                <button
                  className="primaryAction"
                  onClick={addRM}
                >
                  <Plus size={17} />
                  Add RM
                </button>
              </div>

              <p>
                RMs can also appear automatically when
                transaction data is uploaded.
              </p>
            </article>

            <article className="panel rmListPanel">
              <h2>Current RMs</h2>

              <div className="rmList">
                {rms
                  .filter(item => item !== 'All')
                  .map(name => (
                    <div
                      className="rmListItem"
                      key={name}
                    >
                      <span>{name}</span>

                      <button
                        onClick={() =>
                          deleteRM(name)
                        }
                        title="Remove RM"
                      >
                        <Trash2 size={17} />
                      </button>
                    </div>
                  ))}

                {rms.length <= 1 && (
                  <div className="emptyState">
                    No RMs available yet.
                  </div>
                )}
              </div>
            </article>

            <article className="panel passwordPanel">
              <div className="passwordHeading">
                <KeyRound size={23} />
                <div>
                  <h2>Password Reset</h2>
                  <p>
                    Send a password reset link to a user.
                  </p>
                </div>
              </div>

              <div className="addRMRow">
                <input
                  type="email"
                  placeholder="User email address"
                  value={resetEmail}
                  onChange={e =>
                    setResetEmail(e.target.value)
                  }
                />

                <button
                  className="primaryAction"
                  onClick={resetPassword}
                >
                  Send Reset Link
                </button>
              </div>

              {resetMessage && (
                <p className="resetStatus">
                  {resetMessage}
                </p>
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
