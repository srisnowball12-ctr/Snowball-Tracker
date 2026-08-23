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
  UserCheck,
  UserX,
  ArrowLeft,
  KeyRound
} from 'lucide-react'

import './styles.css'
import snowballLogo from './Screenshot 2026-08-18 161429.png'

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
)

/* -------------------- HELPERS -------------------- */

const money = value =>
  new Intl.NumberFormat('en-IN', {
    style: 'currency',
    currency: 'INR',
    maximumFractionDigits: 0
  }).format(Number(value || 0))

const norm = value =>
  String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')

const isoDate = value => {
  if (!value) return null

  if (value instanceof Date) {
    return value.toISOString().slice(0, 10)
  }

  const d = new Date(value)
  return Number.isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10)
}

const getSourceType = row => {
  const value = String(
    row.original_transaction_type ||
    row.transaction_type ||
    ''
  ).toLowerCase()

  if (value.includes('swp')) return 'SWP'
  if (value.includes('stp')) return 'STP'
  if (value.includes('switch')) return 'Switch'
  if (value.includes('red')) return 'Redemption'

  return row.classified_transaction_type || 'Redemption'
}

const getClassification = row =>
  row.classified_transaction_type ||
  row.transaction_type ||
  'Redemption'

function mapExcelRow(row) {
  const lookup = Object.fromEntries(
    Object.entries(row).map(([key, value]) => [norm(key), value])
  )

  const get = (...keys) =>
    keys.map(key => lookup[norm(key)]).find(value => value !== undefined)

  const amountRaw = get('Amount(₹)', 'Amount', 'amount')

  const amount =
    typeof amountRaw === 'number'
      ? amountRaw
      : Number(
          String(amountRaw || '').replace(/[₹,\s]/g, '')
        )

  return {
    rm_name:
      get('Partner/Employee', 'RM', 'RM Name', 'rm_name') || null,

    group_name:
      get('Group', 'group_name') || null,

    investor_name:
      get('Investor', 'Investor Name', 'investor_name') || null,

    transaction_date:
      isoDate(get('Date', 'transaction_date')),

    folio_no:
      String(
        get('Folio No/Demat A/C', 'Folio', 'folio_no') || ''
      ) || null,

    scheme:
      get('Scheme', 'Fund', 'scheme') || null,

    amount:
      Number.isFinite(amount) ? amount : null,

    original_transaction_type:
      get('Type', 'Transaction Type', 'original_transaction_type') || null,

    transaction_type: 'Imported',
    classified_transaction_type: null,
    classification_status: 'Analysed'
  }
}

/* -------------------- CLASSIFICATION ENGINE -------------------- */

/*
  Classification priority:

  1. Explicit SWP / STP / Switch in Excel
  2. Recurring monthly withdrawal pattern = SWP
  3. Otherwise = Redemption
*/

function classifyTransactions(rows) {
  const classified = rows.map(row => ({ ...row }))

  const groups = {}

  classified.forEach((row, index) => {
    const key = [
      row.investor_name || '',
      row.folio_no || '',
      row.scheme || ''
    ].join('|')

    if (!groups[key]) groups[key] = []
    groups[key].push(index)
  })

  Object.values(groups).forEach(indexes => {
    indexes.sort((a, b) =>
      String(classified[a].transaction_date)
        .localeCompare(String(classified[b].transaction_date))
    )

    indexes.forEach(index => {
      const row = classified[index]

      const sourceText = String(
        row.original_transaction_type || ''
      ).toLowerCase()

      /* Explicit classification from source data */

      if (sourceText.includes('swp')) {
        row.classified_transaction_type = 'SWP'
        return
      }

      if (sourceText.includes('stp')) {
        row.classified_transaction_type = 'STP'
        return
      }

      if (sourceText.includes('switch')) {
        row.classified_transaction_type = 'Switch'
        return
      }

      /* Check for recurring SWP pattern */

      const history = indexes
        .filter(i =>
          classified[i].transaction_date &&
          row.transaction_date &&
          classified[i].transaction_date <= row.transaction_date
        )
        .map(i => classified[i])
        .sort((a, b) =>
          String(a.transaction_date).localeCompare(
            String(b.transaction_date)
          )
        )

      let recurringMatches = 0

      for (let i = 1; i < history.length; i++) {
        const previous = history[i - 1]
        const current = history[i]

        const d1 = new Date(previous.transaction_date + 'T00:00:00')
        const d2 = new Date(current.transaction_date + 'T00:00:00')

        const days =
          Math.round((d2 - d1) / (1000 * 60 * 60 * 24))

        const previousAmount = Number(previous.amount || 0)
        const currentAmount = Number(current.amount || 0)

        const difference =
          previousAmount > 0
            ? Math.abs(currentAmount - previousAmount) / previousAmount
            : 1

        /*
          Monthly recurring transaction:
          approximately 20–40 days apart
          and amount reasonably similar
        */

        if (
          days >= 20 &&
          days <= 40 &&
          difference <= 0.25
        ) {
          recurringMatches++
        }
      }

      if (recurringMatches >= 2) {
        row.classified_transaction_type = 'SWP'
      } else {
        row.classified_transaction_type = 'Redemption'
      }
    })
  })

  return classified
}

/* -------------------- MAIN APP -------------------- */

function App() {
  const [session, setSession] = useState(null)

  const [authMode, setAuthMode] = useState('login')
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [newPassword, setNewPassword] = useState('')

  const [page, setPage] = useState('dashboard')

  const [rows, setRows] = useState([])
  const [rms, setRms] = useState([])

  const [rm, setRm] = useState('All')
  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')
  const [period, setPeriod] = useState('YTD')

  const [loading, setLoading] = useState(false)
  const [uploading, setUploading] = useState(false)

  const [error, setError] = useState('')
  const [message, setMessage] = useState('')

  const [newRM, setNewRM] = useState('')

  /* -------------------- AUTH -------------------- */

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session)
    })

    const {
      data: { subscription }
    } = supabase.auth.onAuthStateChange((event, newSession) => {
      setSession(newSession)

      if (event === 'PASSWORD_RECOVERY') {
        setAuthMode('newPassword')
      }
    })

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

  async function sendResetLink(e) {
    e.preventDefault()

    setLoading(true)
    setError('')
    setMessage('')

    const { error } =
      await supabase.auth.resetPasswordForEmail(email, {
        redirectTo: window.location.origin
      })

    if (error) {
      setError(error.message)
    } else {
      setMessage(
        'Password reset link has been sent. Please check your email.'
      )
    }

    setLoading(false)
  }

  async function updatePassword(e) {
    e.preventDefault()

    setLoading(true)
    setError('')

    const { error } =
      await supabase.auth.updateUser({
        password: newPassword
      })

    if (error) {
      setError(error.message)
    } else {
      setMessage(
        'Password updated successfully. You can now use your new password.'
      )

      setAuthMode('login')
      setNewPassword('')
    }

    setLoading(false)
  }

  async function logout() {
    await supabase.auth.signOut()

    setSession(null)
    setPage('dashboard')
    setEmail('')
    setPassword('')
    setError('')
    setMessage('')
    setAuthMode('login')
  }

  /* -------------------- DATA -------------------- */

  async function loadData() {
    setLoading(true)

    const { data, error } =
      await supabase
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
    const { data, error } =
      await supabase
        .from('relationship_managers')
        .select('*')
        .order('name')

    if (error) {
      console.error(error)
      return
    }

    setRms(data || [])
  }

  /* -------------------- FILTERING -------------------- */

  const activeRMs = useMemo(
    () =>
      rms
        .filter(item => item.is_active !== false)
        .map(item => item.name),
    [rms]
  )

  const filtered = useMemo(() => {
    return rows.filter(row => {
      if (rm !== 'All' && row.rm_name !== rm) {
        return false
      }

      const date = row.transaction_date

      if (!date) return false

      if (from && date < from) return false
      if (to && date > to) return false

      if (!from && !to) {
        const now = new Date()
        const dt = new Date(date + 'T00:00:00')

        if (period === 'WTD') {
          const day =
            (now.getDay() + 6) % 7

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

  /* -------------------- DASHBOARD TOTALS -------------------- */

  const totals = useMemo(() => {
    const getTotal = type =>
      filtered
        .filter(
          row =>
            getClassification(row) === type
        )
        .reduce(
          (sum, row) =>
            sum + Number(row.amount || 0),
          0
        )

    return {
      Redemption: getTotal('Redemption'),
      SWP: getTotal('SWP'),
      Switch: getTotal('Switch'),
      STP: getTotal('STP'),
      Investors: new Set(
        filtered
          .map(row => row.investor_name)
          .filter(Boolean)
      ).size,
      Transactions: filtered.length
    }
  }, [filtered])

  /* -------------------- RM-WISE DATA -------------------- */

  const rmSummary = useMemo(() => {
    const summary = {}

    filtered.forEach(row => {
      const name =
        row.rm_name || 'Unassigned'

      summary[name] =
        (summary[name] || 0) +
        Number(row.amount || 0)
    })

    return Object.entries(summary)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8)
  }, [filtered])

  /* -------------------- UPLOAD -------------------- */

  async function uploadFile(e) {
    const file = e.target.files?.[0]

    if (!file) return

    setUploading(true)
    setError('')
    setMessage('Reading Excel file...')

    try {
      const buffer = await file.arrayBuffer()

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
          { defval: null }
        )

      let mapped =
        raw
          .map(mapExcelRow)
          .filter(
            row =>
              row.investor_name &&
              row.transaction_date &&
              row.amount !== null
          )

      if (!mapped.length) {
        throw new Error(
          'No valid transactions were found in this Excel file.'
        )
      }

      setMessage(
        `Analysing ${mapped.length} transactions...`
      )

      /*
        SYSTEM CLASSIFICATION
      */

      mapped =
        classifyTransactions(mapped)

      /*
        Automatically add any new RM
        found in uploaded data.
      */

      const uniqueRMNames =
        [
          ...new Set(
            mapped
              .map(row => row.rm_name)
              .filter(Boolean)
          )
        ]

      for (const name of uniqueRMNames) {
        await supabase
          .from('relationship_managers')
          .upsert(
            {
              name,
              is_active: true
            },
            {
              onConflict: 'name'
            }
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

        if (error) {
          throw error
        }
      }

      setMessage(
        'Done. Transactions analysed and dashboard updated.'
      )

      await loadData()
      await loadRMs()

    } catch (err) {
      setError(
        err.message ||
        'Something went wrong while uploading the file.'
      )

      setMessage('')
    }

    setUploading(false)
    e.target.value = ''
  }

  /* -------------------- RM MANAGEMENT -------------------- */

  async function addRM(e) {
    e.preventDefault()

    const name = newRM.trim()

    if (!name) return

    setError('')

    const { error } =
      await supabase
        .from('relationship_managers')
        .upsert(
          {
            name,
            is_active: true
          },
          {
            onConflict: 'name'
          }
        )

    if (error) {
      setError(error.message)
      return
    }

    setNewRM('')
    await loadRMs()
  }

  async function toggleRMStatus(item) {
    const { error } =
      await supabase
        .from('relationship_managers')
        .update({
          is_active:
            !item.is_active
        })
        .eq('id', item.id)

    if (error) {
      setError(error.message)
      return
    }

    await loadRMs()
  }

  /* -------------------- EXPORTS -------------------- */

  function exportExcel() {
    const output =
      filtered.map(row => ({
        Date: row.transaction_date,
        RM: row.rm_name,
        Investor: row.investor_name,
        Folio: row.folio_no,
        Scheme: row.scheme,
        Amount: row.amount,
        Source: getSourceType(row),
        Classification:
          getClassification(row)
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

    autoTable(doc, {
      startY: 22,
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
          (row.scheme || '').slice(0, 35),
          money(row.amount),
          getSourceType(row),
          getClassification(row)
        ])
    })

    doc.save(
      'snowball-transaction-report.pdf'
    )
  }

  /* -------------------- LOGIN -------------------- */

  if (!session) {
    return (
      <main className="loginPage">
        <section className="loginCard">

          {authMode === 'login' && (
            <>
              <h1>
                Snowball Redemption Tracker
              </h1>

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

              <button
                className="forgotButton"
                onClick={() => {
                  setAuthMode('forgot')
                  setError('')
                  setMessage('')
                }}
              >
                Forgot password?
              </button>
            </>
          )}

          {authMode === 'forgot' && (
            <>
              <button
                className="backButton"
                onClick={() =>
                  setAuthMode('login')
                }
              >
                <ArrowLeft size={16} />
                Back to Login
              </button>

              <h1>
                Reset Password
              </h1>

              <p>
                Enter your registered email address.
              </p>

              <form onSubmit={sendResetLink}>
                <input
                  type="email"
                  placeholder="Registered email"
                  value={email}
                  onChange={e =>
                    setEmail(e.target.value)
                  }
                  required
                />

                <button
                  className="loginButton"
                  disabled={loading}
                >
                  <KeyRound size={17} />
                  Send Reset Link
                </button>
              </form>
            </>
          )}

          {authMode === 'newPassword' && (
            <>
              <h1>
                Create New Password
              </h1>

              <p>
                Enter your new password below.
              </p>

              <form onSubmit={updatePassword}>
                <input
                  type="password"
                  placeholder="New password"
                  value={newPassword}
                  onChange={e =>
                    setNewPassword(e.target.value)
                  }
                  required
                />

                <button
                  className="loginButton"
                  disabled={loading}
                >
                  Update Password
                </button>
              </form>
            </>
          )}

          {error && (
            <div className="authError">
              {error}
            </div>
          )}

          {message && (
            <div className="authMessage">
              {message}
            </div>
          )}

        </section>
      </main>
    )
  }

  /* -------------------- FILTER BAR -------------------- */

  const Filters = ({
    showUpload = false
  }) => (
    <section className="filters">

      <div className="periodButtons">
        {['WTD', 'MTD', 'QTD', 'YTD']
          .map(item => (
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
          setRm(e.target.value)
        }
      >
        <option value="All">
          All RMs
        </option>

        {activeRMs.map(name => (
          <option
            key={name}
            value={name}
          >
            {name}
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
        className="refreshButton"
        onClick={loadData}
        title="Refresh"
      >
        <RefreshCw size={17} />
      </button>

      {showUpload && (
        <label className="uploadButton">
          <Upload size={17} />

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
      )}

      <button onClick={exportExcel}>
        <Download size={17} />
        Excel
      </button>

      <button onClick={exportPDF}>
        <FileText size={17} />
        PDF
      </button>

    </section>
  )

  /* -------------------- DASHBOARD -------------------- */

  function Dashboard() {
    const totalAmount =
      totals.Redemption +
      totals.SWP +
      totals.Switch +
      totals.STP

    const percentages = {
      Redemption:
        totalAmount
          ? (totals.Redemption / totalAmount) * 100
          : 0,

      SWP:
        totalAmount
          ? (totals.SWP / totalAmount) * 100
          : 0,

      Switch:
        totalAmount
          ? (totals.Switch / totalAmount) * 100
          : 0,

      STP:
        totalAmount
          ? (totals.STP / totalAmount) * 100
          : 0
    }

    return (
      <>
        <div className="pageHeader">
          <div>
            <h1>
              Snowball Redemption Tracker
            </h1>

            <p>
              Analyse transactions and monitor redemption activity
            </p>
          </div>
        </div>

        <Filters showUpload={false} />

        <section className="summaryCards">

          <article className="summaryCard redemption">
            <span>Redemption</span>
            <strong>
              {money(totals.Redemption)}
            </strong>
          </article>

          <article className="summaryCard swp">
            <span>SWP</span>
            <strong>
              {money(totals.SWP)}
            </strong>
          </article>

          <article className="summaryCard switch">
            <span>Switch</span>
            <strong>
              {money(totals.Switch)}
            </strong>
          </article>

          <article className="summaryCard stp">
            <span>STP</span>
            <strong>
              {money(totals.STP)}
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

        <section className="dashboardGrid">

          <article className="chartCard">
            <h2>
              Amount by Classification
            </h2>

            <div className="classificationContent">

              <div
                className="donut"
                style={{
                  background: `
                    conic-gradient(
                      #d84b55 0 ${percentages.Redemption}%,
                      #58a48d ${percentages.Redemption}% ${percentages.Redemption + percentages.SWP}%,
                      #5b86c8 ${percentages.Redemption + percentages.SWP}% ${percentages.Redemption + percentages.SWP + percentages.Switch}%,
                      #8b70c9 ${percentages.Redemption + percentages.SWP + percentages.Switch}% 100%
                    )
                  `
                }}
              >
                <div className="donutHole" />
              </div>

              <div className="legend">
                {[
                  ['Redemption', totals.Redemption],
                  ['SWP', totals.SWP],
                  ['Switch', totals.Switch],
                  ['STP', totals.STP]
                ].map(([name, value]) => (
                  <div
                    className="legendRow"
                    key={name}
                  >
                    <span
                      className={
                        `legendDot ${name.toLowerCase()}`
                      }
                    />

                    <span>
                      {name}
                    </span>

                    <b>
                      {money(value)}
                    </b>
                  </div>
                ))}
              </div>

            </div>
          </article>

          <article className="chartCard">
            <h2>
              Transaction Value by RM
            </h2>

            <div className="rmBars">
              {rmSummary.length === 0 && (
                <p className="emptyChart">
                  No transaction data available
                </p>
              )}

              {rmSummary.map(([name, value]) => {
                const max =
                  rmSummary[0]?.[1] || 1

                return (
                  <div
                    className="rmBarRow"
                    key={name}
                  >
                    <span>
                      {name}
                    </span>

                    <div className="barTrack">
                      <div
                        className="barFill"
                        style={{
                          width:
                            `${(value / max) * 100}%`
                        }}
                      />
                    </div>

                    <b>
                      {money(value)}
                    </b>
                  </div>
                )
              })}
            </div>
          </article>

          <article className="chartCard classificationAnalysis">
            <h2>
              Classification Analysis
            </h2>

            {[
              ['Redemption', totals.Redemption],
              ['SWP', totals.SWP],
              ['Switch', totals.Switch],
              ['STP', totals.STP]
            ].map(([name, value]) => {
              const width =
                totalAmount
                  ? (value / totalAmount) * 100
                  : 0

              return (
                <div
                  className="analysisRow"
                  key={name}
                >
                  <span>{name}</span>

                  <div className="analysisTrack">
                    <div
                      className={
                        `analysisFill ${name.toLowerCase()}`
                      }
                      style={{
                        width: `${width}%`
                      }}
                    />
                  </div>

                  <b>
                    {money(value)}
                  </b>
                </div>
              )
            })}
          </article>

        </section>
      </>
    )
  }

  /* -------------------- TRANSACTION DATA -------------------- */

  function TransactionData() {
    return (
      <>
        <div className="pageHeader">
          <div>
            <h1>Transaction Data</h1>

            <p>
              Upload and analyse transaction data
            </p>
          </div>
        </div>

        <Filters showUpload />

        {message && (
          <div className="messageBox">
            {message}
          </div>
        )}

        <section className="dataCard">
          <div className="tableTitle">
            <div>
              <h2>
                Transaction Details
              </h2>

              <span>
                {filtered.length} transactions
              </span>
            </div>
          </div>

          <div className="tableScroll">
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

                      <td>
                        {money(row.amount)}
                      </td>

                      <td>
                        {getSourceType(row)}
                      </td>

                      <td>
                        <span
                          className={
                            `classificationBadge ${getClassification(row).toLowerCase()}`
                          }
                        >
                          {getClassification(row)}
                        </span>
                      </td>
                    </tr>
                  ))}

                {!filtered.length && (
                  <tr>
                    <td
                      colSpan="7"
                      className="emptyTable"
                    >
                      No transactions found.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </section>
      </>
    )
  }

  /* -------------------- MANAGE RMS -------------------- */

  function ManageRMs() {
    return (
      <>
        <div className="pageHeader">
          <div>
            <h1>
              Manage Relationship Managers
            </h1>

            <p>
              Manage RMs in your organisation
            </p>
          </div>
        </div>

        <section className="rmAddCard">
          <h2>Add New RM</h2>

          <form onSubmit={addRM}>
            <input
              placeholder="Enter RM name"
              value={newRM}
              onChange={e =>
                setNewRM(e.target.value)
              }
            />

            <button
              type="submit"
              className="primaryButton"
            >
              <Plus size={17} />
              Add RM
            </button>
          </form>
        </section>

        <section className="rmListCard">
          <h2>
            Current RMs
          </h2>

          <div className="rmList">
            {rms.map(item => (
              <div
                className="rmRow"
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
                      ? 'inactiveButton'
                      : 'activateButton'
                  }
                  onClick={() =>
                    toggleRMStatus(item)
                  }
                >
                  {item.is_active ? (
                    <>
                      <UserX size={16} />
                      Make Inactive
                    </>
                  ) : (
                    <>
                      <UserCheck size={16} />
                      Make Active
                    </>
                  )}
                </button>
              </div>
            ))}

            {!rms.length && (
              <p className="emptyChart">
                No RMs added yet.
              </p>
            )}
          </div>
        </section>
      </>
    )
  }

  /* -------------------- APP LAYOUT -------------------- */

  return (
    <div className="appShell">

      <aside className="sidebar">

        <div className="brand">
          <img
            src={snowballLogo}
            alt="Snowball Financial Services"
          />
        </div>

        <nav>

          <button
            className={
              page === 'dashboard'
                ? 'navItem activeNav'
                : 'navItem'
            }
            onClick={() =>
              setPage('dashboard')
            }
          >
            <LayoutDashboard size={19} />
            Dashboard
          </button>

          <button
            className={
              page === 'transactions'
                ? 'navItem activeNav'
                : 'navItem'
            }
            onClick={() =>
              setPage('transactions')
            }
          >
            <Table2 size={19} />
            Transaction Data
          </button>

          <button
            className={
              page === 'rms'
                ? 'navItem activeNav'
                : 'navItem'
            }
            onClick={() =>
              setPage('rms')
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

      <main className="mainContent">

        <header className="topBar">

          <div />

          <button
            className="logoutButton"
            onClick={logout}
          >
            <LogOut size={17} />
            Logout
          </button>

        </header>

        {error && (
          <div className="errorBox">
            {error}
          </div>
        )}

        {page === 'dashboard' && (
          <Dashboard />
        )}

        {page === 'transactions' && (
          <TransactionData />
        )}

        {page === 'rms' && (
          <ManageRMs />
        )}

      </main>
    </div>
  )
}

createRoot(
  document.getElementById('root')
).render(<App />)
