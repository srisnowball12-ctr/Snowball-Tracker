import React, { useEffect, useMemo, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'
import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'

import {
  LayoutDashboard,
  TableProperties,
  Settings,
  LogOut,
  Upload,
  Download,
  FileText,
  RefreshCw,
  Users,
  Plus,
  UserCheck,
  UserX,
  ChevronRight
} from 'lucide-react'

import './styles.css'
import logo from './logo.png'

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
)

const money = (value) =>
  new Intl.NumberFormat('en-IN', {
    style: 'currency',
    currency: 'INR',
    maximumFractionDigits: 0
  }).format(Number(value || 0))

const numberFormat = (value) =>
  new Intl.NumberFormat('en-IN', {
    maximumFractionDigits: 0
  }).format(Number(value || 0))

const normalize = (value) =>
  String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')

const formatDate = (value) => {
  if (!value) return '-'

  const date = new Date(`${value}T00:00:00`)

  if (Number.isNaN(date.getTime())) return value

  return date.toLocaleDateString('en-IN', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  })
}

const toISODate = (value) => {
  if (!value) return null

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10)
  }

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value)

    if (parsed) {
      const month = String(parsed.m).padStart(2, '0')
      const day = String(parsed.d).padStart(2, '0')

      return `${parsed.y}-${month}-${day}`
    }
  }

  const date = new Date(value)

  if (Number.isNaN(date.getTime())) return null

  return date.toISOString().slice(0, 10)
}

function mapExcelRow(row) {
  const lookup = Object.fromEntries(
    Object.entries(row).map(([key, value]) => [normalize(key), value])
  )

  const get = (...keys) => {
    for (const key of keys) {
      const value = lookup[normalize(key)]

      if (value !== undefined && value !== null && value !== '') {
        return value
      }
    }

    return null
  }

  const rawAmount = get(
    'Amount(₹)',
    'Amount',
    'Transaction Amount',
    'Value'
  )

  const amount =
    typeof rawAmount === 'number'
      ? rawAmount
      : Number(
          String(rawAmount || '').replace(/[₹,\s]/g, '')
        )

  const rawType = String(
    get(
      'Type',
      'Transaction Type',
      'TransactionType',
      'Nature'
    ) || ''
  ).trim()

  return {
    rm_name:
      get(
        'Partner/Employee',
        'Partner Employee',
        'RM',
        'RM Name',
        'Employee'
      ) || null,

    group_name:
      get('Group', 'Group Name') || null,

    investor_name:
      get(
        'Investor',
        'Investor Name',
        'Client',
        'Client Name'
      ) || null,

    transaction_date:
      toISODate(
        get(
          'Date',
          'Transaction Date',
          'Txn Date'
        )
      ),

    folio_no: String(
      get(
        'Folio No/Demat A/C',
        'Folio No',
        'Folio',
        'Demat A/C'
      ) || ''
    ) || null,

    scheme:
      get(
        'Scheme',
        'Scheme Name',
        'Fund'
      ) || null,

    amount: Number.isFinite(amount) ? amount : null,

    transaction_type: 'Imported',

    original_transaction_type: rawType || null,

    classified_transaction_type: null,

    classification_status: 'Needs Review',

    classification_reason: null
  }
}

/*
  Automatic classification.

  Priority:
  1. Explicit transaction type in Excel
  2. Recurring same investor/folio/scheme pattern = SWP
  3. Switch keywords = Switch
  4. STP keywords = STP
  5. Remaining transactions = Redemption
*/

function classifyTransactions(rows) {
  const prepared = rows.map((row) => ({
    ...row,
    classified_transaction_type: null
  }))

  const groups = {}

  prepared.forEach((row, index) => {
    const investor = normalize(row.investor_name)
    const folio = normalize(row.folio_no)
    const scheme = normalize(row.scheme)

    const key = `${investor}|${folio}|${scheme}`

    if (!groups[key]) groups[key] = []

    groups[key].push(index)
  })

  Object.values(groups).forEach((indexes) => {
    indexes.sort((a, b) => {
      return new Date(prepared[a].transaction_date) -
        new Date(prepared[b].transaction_date)
    })

    const recurringIndexes = new Set()

    for (let i = 1; i < indexes.length; i++) {
      const previous = prepared[indexes[i - 1]]
      const current = prepared[indexes[i]]

      const previousDate = new Date(
        `${previous.transaction_date}T00:00:00`
      )

      const currentDate = new Date(
        `${current.transaction_date}T00:00:00`
      )

      const difference =
        (currentDate - previousDate) /
        (1000 * 60 * 60 * 24)

      const previousAmount = Number(previous.amount || 0)
      const currentAmount = Number(current.amount || 0)

      const amountDifference =
        previousAmount === 0
          ? 100
          : Math.abs(currentAmount - previousAmount) /
            previousAmount *
            100

      /*
        Monthly / approximately recurring withdrawal pattern.
      */
      if (
        difference >= 20 &&
        difference <= 40 &&
        amountDifference <= 20
      ) {
        recurringIndexes.add(indexes[i - 1])
        recurringIndexes.add(indexes[i])
      }
    }

    recurringIndexes.forEach((index) => {
      if (!prepared[index].classified_transaction_type) {
        prepared[index].classified_transaction_type = 'SWP'
      }
    })
  })

  return prepared.map((row) => {
    const source = `${row.original_transaction_type || ''} ${
      row.scheme || ''
    }`.toLowerCase()

    let classification = row.classified_transaction_type

    if (
      source.includes('stp') ||
      source.includes('systematic transfer')
    ) {
      classification = 'STP'
    } else if (
      source.includes('switch')
    ) {
      classification = 'Switch'
    } else if (
      source.includes('swp') ||
      source.includes('systematic withdrawal')
    ) {
      classification = 'SWP'
    }

    if (!classification) {
      classification = 'Redemption'
    }

    return {
      ...row,
      classified_transaction_type: classification,
      classification_status: 'Classified'
    }
  })
}

function App() {
  const [session, setSession] = useState(null)

  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')

  const [loading, setLoading] = useState(false)
  const [uploading, setUploading] = useState(false)

  const [error, setError] = useState('')
  const [message, setMessage] = useState('')

  const [page, setPage] = useState('dashboard')

  const [rows, setRows] = useState([])
  const [rms, setRms] = useState([])

  const [rmFilter, setRmFilter] = useState('All')
  const [period, setPeriod] = useState('YTD')

  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')

  const [newRM, setNewRM] = useState('')

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      setSession(data.session)
    })

    const {
      data: { subscription }
    } = supabase.auth.onAuthStateChange(
      (_event, currentSession) => {
        setSession(currentSession)
      }
    )

    return () => subscription.unsubscribe()
  }, [])

  useEffect(() => {
    if (session) {
      loadAllData()
    }
  }, [session])

  async function login(event) {
    event.preventDefault()

    setLoading(true)
    setError('')

    try {
      const { error: loginError } =
        await supabase.auth.signInWithPassword({
          email,
          password
        })

      if (loginError) {
        throw loginError
      }
    } catch (err) {
      setError(err.message || 'Unable to login.')
    } finally {
      setLoading(false)
    }
  }

  async function logout() {
    try {
      setLoading(true)
      setError('')

      await supabase.auth.signOut()

      setSession(null)
      setEmail('')
      setPassword('')

      /*
        Remove any locally cached Supabase session.
      */
      Object.keys(localStorage).forEach((key) => {
        if (
          key.startsWith('sb-') ||
          key.includes('supabase')
        ) {
          localStorage.removeItem(key)
        }
      })

      window.location.reload()
    } catch (err) {
      setError(err.message || 'Unable to logout.')
    } finally {
      setLoading(false)
    }
  }

  async function loadTransactions() {
    const { data, error: transactionError } =
      await supabase
        .from('transactions')
        .select('*')
        .order('transaction_date', {
          ascending: false
        })

    if (transactionError) {
      throw transactionError
    }

    setRows(data || [])
  }

  async function loadRMs() {
    /*
      We first try to load the dedicated RMS table.
      If your existing database table is named differently,
      this can be adjusted later.
    */

    const { data, error: rmError } =
      await supabase
        .from('rms')
        .select('*')
        .order('rm_name', {
          ascending: true
        })

    if (!rmError) {
      setRms(data || [])
      return
    }

    /*
      Fallback: derive RM names from transactions.
    */

    const { data: transactionData } =
      await supabase
        .from('transactions')
        .select('rm_name')

    const names = Array.from(
      new Set(
        (transactionData || [])
          .map((item) => item.rm_name)
          .filter(Boolean)
      )
    ).sort()

    setRms(
      names.map((rm_name) => ({
        id: rm_name,
        rm_name,
        is_active: true,
        fallback: true
      }))
    )
  }

  async function loadAllData() {
    setLoading(true)
    setError('')

    try {
      await Promise.all([
        loadTransactions(),
        loadRMs()
      ])
    } catch (err) {
      setError(err.message || 'Unable to load data.')
    } finally {
      setLoading(false)
    }
  }

  const activeRMs = useMemo(() => {
    return rms.filter(
      (item) => item.is_active !== false
    )
  }, [rms])

  const rmNames = useMemo(() => {
    return [
      'All',
      ...activeRMs
        .map((item) => item.rm_name)
        .filter(Boolean)
        .sort()
    ]
  }, [activeRMs])

  const filteredRows = useMemo(() => {
    return rows.filter((row) => {
      if (
        rmFilter !== 'All' &&
        row.rm_name !== rmFilter
      ) {
        return false
      }

      const transactionDate =
        row.transaction_date

      if (!transactionDate) return false

      if (from && transactionDate < from) {
        return false
      }

      if (to && transactionDate > to) {
        return false
      }

      if (!from && !to) {
        const now = new Date()

        const date = new Date(
          `${transactionDate}T00:00:00`
        )

        if (period === 'WTD') {
          const day =
            (now.getDay() + 6) % 7

          const start = new Date(now)

          start.setDate(
            now.getDate() - day
          )

          start.setHours(0, 0, 0, 0)

          if (date < start) return false
        }

        if (
          period === 'MTD' &&
          (
            date.getMonth() !== now.getMonth() ||
            date.getFullYear() !== now.getFullYear()
          )
        ) {
          return false
        }

        if (period === 'QTD') {
          const currentQuarter =
            Math.floor(now.getMonth() / 3)

          const rowQuarter =
            Math.floor(date.getMonth() / 3)

          if (
            date.getFullYear() !== now.getFullYear() ||
            rowQuarter !== currentQuarter
          ) {
            return false
          }
        }

        if (
          period === 'YTD' &&
          date.getFullYear() !== now.getFullYear()
        ) {
          return false
        }
      }

      return true
    })
  }, [
    rows,
    rmFilter,
    from,
    to,
    period
  ])

  const totals = useMemo(() => {
    const getTotal = (type) =>
      filteredRows
        .filter(
          (row) =>
            row.classified_transaction_type === type
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
        filteredRows
          .map((row) => row.investor_name)
          .filter(Boolean)
      ).size,
      Transactions: filteredRows.length
    }
  }, [filteredRows])

  const classificationData = useMemo(() => {
    const values = [
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

    return values
  }, [totals])

  const rmChartData = useMemo(() => {
    const totalsByRM = {}

    filteredRows.forEach((row) => {
      const name = row.rm_name || 'Others'

      totalsByRM[name] =
        (totalsByRM[name] || 0) +
        Number(row.amount || 0)
    })

    return Object.entries(totalsByRM)
      .map(([name, value]) => ({
        name,
        value
      }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 6)
  }, [filteredRows])

  const monthlyData = useMemo(() => {
    const months = {}

    filteredRows.forEach((row) => {
      if (!row.transaction_date) return

      const key =
        row.transaction_date.slice(0, 7)

      if (!months[key]) {
        months[key] = {
          Redemption: 0,
          SWP: 0,
          Switch: 0,
          STP: 0
        }
      }

      const type =
        row.classified_transaction_type

      if (months[key] && months[key][type] !== undefined) {
        months[key][type] +=
          Number(row.amount || 0)
      }
    })

    return Object.entries(months)
      .sort(([a], [b]) =>
        a.localeCompare(b)
      )
      .slice(-6)
      .map(([month, values]) => ({
        month,
        ...values
      }))
  }, [filteredRows])

  async function uploadFile(event) {
    const file = event.target.files?.[0]

    if (!file) return

    setUploading(true)
    setError('')
    setMessage('Reading Excel file...')

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
          {
            defval: null,
            raw: true
          }
        )

      const mapped = raw
        .map(mapExcelRow)
        .filter(
          (row) =>
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
        Bring existing data into the analysis so recurring
        SWP patterns can also be identified against history.
      */

      const existingForAnalysis =
        rows.map((row) => ({
          ...row
        }))

      const analysed = classifyTransactions([
        ...existingForAnalysis,
        ...mapped
      ])

      const classifiedUpload =
        analysed.slice(existingForAnalysis.length)

      setMessage(
        `Uploading ${classifiedUpload.length} classified transactions...`
      )

      for (
        let index = 0;
        index < classifiedUpload.length;
        index += 500
      ) {
        const batch =
          classifiedUpload.slice(
            index,
            index + 500
          )

        const {
          error: uploadError
        } =
          await supabase
            .from('transactions')
            .insert(batch)

        if (uploadError) {
          throw uploadError
        }
      }

      setMessage(
        'Excel uploaded and transactions classified successfully.'
      )

      await loadAllData()
    } catch (err) {
      setError(
        err.message ||
          'Unable to upload the Excel file.'
      )

      setMessage('')
    } finally {
      setUploading(false)
      event.target.value = ''
    }
  }

  async function addRM() {
    const name = newRM.trim()

    if (!name) return

    setError('')
    setMessage('')

    try {
      const existing = rms.find(
        (item) =>
          normalize(item.rm_name) ===
          normalize(name)
      )

      if (existing) {
        throw new Error(
          'This RM already exists.'
        )
      }

      const {
        error: insertError
      } =
        await supabase
          .from('rms')
          .insert({
            rm_name: name,
            is_active: true
          })

      if (insertError) {
        throw insertError
      }

      setNewRM('')
      setMessage('RM added successfully.')

      await loadRMs()
    } catch (err) {
      setError(err.message)
    }
  }

  async function toggleRM(item) {
    if (item.fallback) {
      setError(
        'The RM management table is not yet available in Supabase.'
      )

      return
    }

    try {
      const {
        error: updateError
      } =
        await supabase
          .from('rms')
          .update({
            is_active:
              item.is_active === false
                ? true
                : false
          })
          .eq('id', item.id)

      if (updateError) {
        throw updateError
      }

      await loadRMs()

      setMessage(
        item.is_active === false
          ? 'RM activated successfully.'
          : 'RM marked inactive successfully.'
      )
    } catch (err) {
      setError(err.message)
    }
  }

  function exportExcel() {
    const output = filteredRows.map(
      (row) => ({
        Date: row.transaction_date,
        RM: row.rm_name,
        Investor: row.investor_name,
        Folio: row.folio_no,
        Scheme: row.scheme,
        Amount: row.amount,
        Source:
          row.classified_transaction_type,
        Classification:
          row.classified_transaction_type
      })
    )

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
    const doc = new jsPDF({
      orientation: 'landscape'
    })

    doc.setFontSize(16)

    doc.text(
      'Snowball Redemption Tracker',
      14,
      15
    )

    doc.setFontSize(9)

    doc.text(
      `Period: ${
        from || to
          ? `${from || ''} to ${to || ''}`
          : period
      }`,
      14,
      22
    )

    autoTable(doc, {
      startY: 28,

      head: [[
        'Date',
        'RM',
        'Investor',
        'Scheme',
        'Amount',
        'Classification'
      ]],

      body: filteredRows.map(
        (row) => [
          formatDate(row.transaction_date),
          row.rm_name || '-',
          row.investor_name || '-',
          String(row.scheme || '').slice(0, 35),
          money(row.amount),
          row.classified_transaction_type || '-'
        ]
      )
    })

    doc.save(
      'snowball-transaction-report.pdf'
    )
  }

  const maxRMValue =
    Math.max(
      ...rmChartData.map(
        (item) => item.value
      ),
      1
    )

  const maxClassification =
    Math.max(
      ...classificationData.map(
        (item) => item.value
      ),
      1
    )

  if (!session) {
    return (
      <main className="login-page">
        <section className="login-card">
          <img
            src={logo}
            className="login-logo"
            alt="Snowball Financial Services"
          />

          <h1>Snowball Redemption Tracker</h1>

          <p>
            Analyse transactions and monitor
            redemption activity
          </p>

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
              className="login-button"
              disabled={loading}
            >
              {loading
                ? 'Signing in...'
                : 'Login'}
            </button>

            <button
              type="button"
              className="forgot-password"
              onClick={() =>
                setMessage(
                  'Please contact the administrator to reset your password.'
                )
              }
            >
              Forgot / Reset Password?
            </button>

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
          </form>
        </section>
      </main>
    )
  }

  return (
    <main className="app-shell">

      <aside className="sidebar">
        <div className="brand">
          <img
            src={logo}
            alt="Snowball Financial Services"
          />
        </div>

        <nav className="navigation">

          <button
            className={
              page === 'dashboard'
                ? 'nav-item active'
                : 'nav-item'
            }
            onClick={() =>
              setPage('dashboard')
            }
          >
            <LayoutDashboard size={18} />
            <span>Dashboard</span>
          </button>

          <button
            className={
              page === 'transactions'
                ? 'nav-item active'
                : 'nav-item'
            }
            onClick={() =>
              setPage('transactions')
            }
          >
            <TableProperties size={18} />
            <span>Transaction Data</span>
          </button>

          <button
            className={
              page === 'settings'
                ? 'nav-item active'
                : 'nav-item'
            }
            onClick={() =>
              setPage('settings')
            }
          >
            <Settings size={18} />
            <span>Settings</span>
          </button>

        </nav>

        <div className="sidebar-bottom">
          <button
            className="logout-button"
            onClick={logout}
          >
            <LogOut size={18} />
            Logout
          </button>
        </div>
      </aside>

      <section className="main-content">

        <header className="top-header">
          <div>
            <h1>Snowball Redemption Tracker</h1>

            <p>
              Analyse transactions and monitor
              redemption activity
            </p>
          </div>

          <div className="header-user">
            <span>
              {session.user?.email}
            </span>

            <div className="avatar">
              {String(
                session.user?.email || 'S'
              )
                .charAt(0)
                .toUpperCase()}
            </div>
          </div>
        </header>

        {error && (
          <div className="error global-message">
            {error}
          </div>
        )}

        {message && (
          <div className="message global-message">
            {message}
          </div>
        )}

        <section className="filter-bar">

          <div className="period-buttons">
            {[
              'WTD',
              'MTD',
              'QTD',
              'YTD'
            ].map((item) => (
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
            value={rmFilter}
            onChange={(event) =>
              setRmFilter(
                event.target.value
              )
            }
          >
            {rmNames.map((item) => (
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

          <button
            className="refresh-button"
            onClick={loadAllData}
            title="Refresh data"
          >
            <RefreshCw size={17} />
          </button>

        </section>

        {page === 'dashboard' && (
          <>

            <section className="metric-grid">

              <article className="metric-card redemption-card">
                <span>Redemption</span>
                <strong>
                  {money(totals.Redemption)}
                </strong>
              </article>

              <article className="metric-card swp-card">
                <span>SWP</span>
                <strong>
                  {money(totals.SWP)}
                </strong>
              </article>

              <article className="metric-card switch-card">
                <span>Switch</span>
                <strong>
                  {money(totals.Switch)}
                </strong>
              </article>

              <article className="metric-card stp-card">
                <span>STP</span>
                <strong>
                  {money(totals.STP)}
                </strong>
              </article>

              <article className="metric-card investor-card">
                <span>Investors</span>
                <strong>
                  {numberFormat(
                    totals.Investors
                  )}
                </strong>
              </article>

              <article className="metric-card transaction-card">
                <span>Transactions</span>
                <strong>
                  {numberFormat(
                    totals.Transactions
                  )}
                </strong>
              </article>

            </section>

            <section className="dashboard-grid">

              <article className="chart-card classification-card">
                <h3>
                  Amount by Classification
                </h3>

                <div className="classification-content">

                  <div className="donut-wrap">
                    <div className="donut">
                      <div className="donut-center">
                        <strong>
                          {numberFormat(
                            totals.Transactions
                          )}
                        </strong>

                        <span>
                          Transactions
                        </span>
                      </div>
                    </div>
                  </div>

                  <div className="legend">
                    {classificationData.map(
                      (item) => (
                        <div
                          className="legend-row"
                          key={item.label}
                        >
                          <div className="legend-name">
                            <span
                              className={`dot ${item.className}`}
                            />

                            {item.label}
                          </div>

                          <strong>
                            {money(item.value)}
                          </strong>
                        </div>
                      )
                    )}
                  </div>

                </div>
              </article>

              <article className="chart-card monthly-card">
                <h3>
                  Monthly Trend
                </h3>

                <div className="trend-chart">
                  {monthlyData.length === 0 ? (
                    <div className="empty-chart">
                      No transaction data available
                    </div>
                  ) : (
                    monthlyData.map(
                      (item) => {
                        const total =
                          item.Redemption +
                          item.SWP +
                          item.Switch +
                          item.STP

                        const max =
                          Math.max(
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
                            className="trend-column"
                            key={item.month}
                          >
                            <div className="trend-bars">
                              <div
                                className="trend-bar redemption"
                                style={{
                                  height: `${(
                                    item.Redemption /
                                    max
                                  ) * 100}%`
                                }}
                              />

                              <div
                                className="trend-bar swp"
                                style={{
                                  height: `${(
                                    item.SWP /
                                    max
                                  ) * 100}%`
                                }}
                              />

                              <div
                                className="trend-bar switch"
                                style={{
                                  height: `${(
                                    item.Switch /
                                    max
                                  ) * 100}%`
                                }}
                              />

                              <div
                                className="trend-bar stp"
                                style={{
                                  height: `${(
                                    item.STP /
                                    max
                                  ) * 100}%`
                                }}
                              />
                            </div>

                            <span>
                              {item.month.slice(5)}
                            </span>
                          </div>
                        )
                      }
                    )
                  )}
                </div>
              </article>

              <article className="chart-card rm-card">
                <h3>
                  Transactions by RM
                </h3>

                <div className="horizontal-chart">
                  {rmChartData.length === 0 ? (
                    <div className="empty-chart">
                      No RM data available
                    </div>
                  ) : (
                    rmChartData.map(
                      (item) => (
                        <div
                          className="rm-bar-row"
                          key={item.name}
                        >
                          <span className="rm-chart-name">
                            {item.name}
                          </span>

                          <div className="rm-bar-track">
                            <div
                              className="rm-bar-fill"
                              style={{
                                width: `${(
                                  item.value /
                                  maxRMValue
                                ) * 100}%`
                              }}
                            />
                          </div>

                          <strong>
                            {money(item.value)}
                          </strong>
                        </div>
                      )
                    )
                  )}
                </div>
              </article>

              <article className="chart-card analysis-card">
                <h3>
                  Classification Analysis
                </h3>

                <div className="classification-bars">
                  {classificationData.map(
                    (item) => (
                      <div
                        className="analysis-row"
                        key={item.label}
                      >
                        <div className="analysis-label">
                          <span
                            className={`dot ${item.className}`}
                          />

                          {item.label}
                        </div>

                        <div className="analysis-track">
                          <div
                            className={`analysis-fill ${item.className}`}
                            style={{
                              width: `${(
                                item.value /
                                maxClassification
                              ) * 100}%`
                            }}
                          />
                        </div>

                        <strong>
                          {money(item.value)}
                        </strong>
                      </div>
                    )
                  )}
                </div>
              </article>

            </section>

            <section className="recent-section">
              <div className="section-heading">
                <div>
                  <h2>
                    Recent Transactions
                  </h2>

                  <p>
                    Latest classified transaction activity
                  </p>
                </div>

                <button
                  onClick={() =>
                    setPage('transactions')
                  }
                >
                  View All Transactions
                  <ChevronRight size={16} />
                </button>
              </div>

              <div className="recent-table-wrap">
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
                    {filteredRows
                      .slice(0, 8)
                      .map((row) => (
                        <tr key={row.id}>
                          <td>
                            {formatDate(
                              row.transaction_date
                            )}
                          </td>

                          <td>
                            {row.rm_name || '-'}
                          </td>

                          <td>
                            {row.investor_name || '-'}
                          </td>

                          <td>
                            {row.scheme || '-'}
                          </td>

                          <td className="amount-cell">
                            {money(row.amount)}
                          </td>

                          <td>
                            {row.classified_transaction_type || '-'}
                          </td>

                          <td>
                            <span
                              className={`classification-pill ${String(
                                row.classified_transaction_type || ''
                              ).toLowerCase()}`}
                            >
                              {row.classified_transaction_type || '-'}
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
          <section className="transaction-page">

            <div className="page-heading">
              <div>
                <h2>
                  Transaction Data
                </h2>

                <p>
                  Upload, review and export transaction data
                </p>
              </div>

              <div className="transaction-actions">

                <button className="upload-button">
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
            </div>

            <div className="table-card">

              <div className="table-title">
                Transactions (
                {filteredRows.length})
              </div>

              <div className="full-table-wrap">
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
                    {filteredRows
                      .slice(0, 1000)
                      .map((row) => (
                        <tr key={row.id}>
                          <td>
                            {formatDate(
                              row.transaction_date
                            )}
                          </td>

                          <td>
                            {row.rm_name || '-'}
                          </td>

                          <td>
                            {row.investor_name || '-'}
                          </td>

                          <td>
                            {row.folio_no || '-'}
                          </td>

                          <td>
                            {row.scheme || '-'}
                          </td>

                          <td className="amount-cell">
                            {money(row.amount)}
                          </td>

                          <td>
                            {row.classified_transaction_type || '-'}
                          </td>

                          <td>
                            <span
                              className={`classification-pill ${String(
                                row.classified_transaction_type || ''
                              ).toLowerCase()}`}
                            >
                              {row.classified_transaction_type || '-'}
                            </span>
                          </td>
                        </tr>
                      ))}
                  </tbody>
                </table>
              </div>
            </div>

          </section>
        )}

        {page === 'settings' && (
          <section className="settings-page">

            <div className="page-heading">
              <div>
                <h2>
                  Settings
                </h2>

                <p>
                  Add and manage Relationship Managers
                </p>
              </div>
            </div>

            <div className="settings-card">

              <h3>
                Manage RMs
              </h3>

              <p className="settings-note">
                RMs are never deleted. When an RM leaves the organisation, mark the RM inactive. Historical transaction data remains unchanged.
              </p>

              <div className="add-rm-row">
                <input
                  value={newRM}
                  placeholder="Enter RM name"
                  onChange={(event) =>
                    setNewRM(
                      event.target.value
                    )
                  }
                  onKeyDown={(event) => {
                    if (
                      event.key === 'Enter'
                    ) {
                      addRM()
                    }
                  }}
                />

                <button
                  className="add-rm-button"
                  onClick={addRM}
                >
                  <Plus size={17} />
                  Add RM
                </button>
              </div>

              <div className="rm-list">

                {rms.length === 0 ? (
                  <div className="empty-rm">
                    No RMs have been added yet.
                  </div>
                ) : (
                  rms.map((item) => (
                    <div
                      className="rm-row"
                      key={item.id}
                    >
                      <div className="rm-details">
                        <Users size={19} />

                        <strong>
                          {item.rm_name}
                        </strong>

                        <span
                          className={
                            item.is_active === false
                              ? 'status inactive'
                              : 'status active'
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
                            ? 'activate-button'
                            : 'inactive-button'
                        }
                        onClick={() =>
                          toggleRM(item)
                        }
                      >
                        {item.is_active === false ? (
                          <>
                            <UserCheck size={16} />
                            Activate
                          </>
                        ) : (
                          <>
                            <UserX size={16} />
                            Mark Inactive
                          </>
                        )}
                      </button>
                    </div>
                  ))
                )}

              </div>
            </div>
          </section>
        )}

      </section>
    </main>
  )
}

createRoot(
  document.getElementById('root')
).render(<App />)
