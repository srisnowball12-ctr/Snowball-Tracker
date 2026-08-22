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
  Settings,
  Upload,
  Download,
  FileText,
  LogOut,
  RefreshCw,
  Plus,
  Pencil,
  UserCheck,
  UserX,
  ShieldCheck,
  X,
  ChevronRight
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

  const d = new Date(v)
  return isNaN(d) ? null : d.toISOString().slice(0, 10)
}

const norm = s =>
  String(s || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')

function mapRow(row) {
  const lookup = Object.fromEntries(
    Object.entries(row).map(([k, v]) => [norm(k), v])
  )

  const get = (...keys) =>
    keys.map(k => lookup[norm(k)]).find(v => v !== undefined)

  const amountRaw = get('Amount(₹)', 'Amount', 'amount')

  const amount =
    typeof amountRaw === 'number'
      ? amountRaw
      : Number(String(amountRaw || '').replace(/[₹,\s]/g, ''))

  return {
    rm_name: get(
      'Partner/Employee',
      'Partner Employee',
      'RM',
      'rm_name'
    ) || null,

    group_name: get('Group', 'group_name') || null,

    investor_name: get(
      'Investor',
      'investor_name'
    ) || null,

    transaction_date: iso(
      get('Date', 'transaction_date')
    ),

    folio_no: String(
      get(
        'Folio No/Demat A/C',
        'Folio No',
        'Folio',
        'folio_no'
      ) || ''
    ) || null,

    scheme: get(
      'Scheme',
      'Fund',
      'scheme'
    ) || null,

    amount: Number.isFinite(amount)
      ? amount
      : null,

    transaction_type: get(
      'Type',
      'transaction_type'
    ) || 'Imported',

    original_transaction_type: get(
      'Type',
      'original_transaction_type'
    ) || null,

    classified_transaction_type: null,

    classification_status: 'Pending',

    classification_reason: null
  }
}

function App() {
  const [session, setSession] = useState(null)

  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [error, setError] = useState('')
  const [message, setMessage] = useState('')
  const [loading, setLoading] = useState(false)

  const [activePage, setActivePage] = useState('dashboard')

  const [rows, setRows] = useState([])
  const [rms, setRms] = useState([])
  const [admins, setAdmins] = useState([])

  const [rm, setRm] = useState('All')
  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')
  const [period, setPeriod] = useState('YTD')

  const [uploading, setUploading] = useState(false)

  const [newRM, setNewRM] = useState('')
  const [editingRM, setEditingRM] = useState(null)
  const [editRMName, setEditRMName] = useState('')

  const [newAdminName, setNewAdminName] = useState('')
  const [newAdminEmail, setNewAdminEmail] = useState('')

  const [showForgot, setShowForgot] = useState(false)
  const [resetEmail, setResetEmail] = useState('')

  useEffect(() => {
    supabase.auth
      .getSession()
      .then(({ data }) => {
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
      loadAllData()
    }
  }, [session])

  async function loadAllData() {
    setLoading(true)
    setError('')

    await Promise.all([
      loadTransactions(),
      loadRMs(),
      loadAdmins()
    ])

    setLoading(false)
  }

  async function loadTransactions() {
    const { data, error } = await supabase
      .from('transactions')
      .select('*')
      .order('transaction_date', {
        ascending: false
      })

    if (error) {
      setError(error.message)
      return
    }

    setRows(data || [])
  }

  async function loadRMs() {
    const { data, error } = await supabase
      .from('relationship_managers')
      .select('*')
      .order('rm_name')

    if (error) {
      console.log('RM table error:', error.message)
      return
    }

    setRms(data || [])
  }

  async function loadAdmins() {
    const { data, error } = await supabase
      .from('admin_users')
      .select('*')
      .order('email')

    if (error) {
      console.log('Admin table error:', error.message)
      return
    }

    setAdmins(data || [])
  }

  async function login(e) {
    e.preventDefault()

    setLoading(true)
    setError('')
    setMessage('')

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

  async function forgotPassword(e) {
    e.preventDefault()

    if (!resetEmail) {
      setError('Please enter your email address.')
      return
    }

    setLoading(true)
    setError('')
    setMessage('')

    const { error } =
      await supabase.auth.resetPasswordForEmail(
        resetEmail,
        {
          redirectTo: window.location.origin
        }
      )

    if (error) {
      setError(error.message)
    } else {
      setMessage(
        'Password reset instructions have been sent to your email.'
      )
      setResetEmail('')
    }

    setLoading(false)
  }

  async function logout() {
    await supabase.auth.signOut()
    setActivePage('dashboard')
  }

  const activeRMNames = useMemo(() => {
    return rms
      .filter(x => x.status !== 'Inactive')
      .map(x => x.rm_name)
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
    const getTotal = type =>
      filtered
        .filter(
          x =>
            String(
              x.classified_transaction_type || ''
            ).toLowerCase() === type.toLowerCase()
        )
        .reduce(
          (sum, x) =>
            sum + Number(x.amount || 0),
          0
        )

    return {
      Redemption: getTotal('Redemption'),
      SWP: getTotal('SWP'),
      Switch: getTotal('Switch'),
      STP: getTotal('STP'),
      Investors: new Set(
        filtered
          .map(x => x.investor_name)
          .filter(Boolean)
      ).size,
      Transactions: filtered.length
    }
  }, [filtered])

  const rmChart = useMemo(() => {
    const result = {}

    filtered.forEach(x => {
      const name = x.rm_name || 'Others'

      result[name] =
        (result[name] || 0) +
        Number(x.amount || 0)
    })

    return Object.entries(result)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 7)
  }, [filtered])

  const monthlyData = useMemo(() => {
    const result = {}

    filtered.forEach(x => {
      if (!x.transaction_date) {
        return
      }

      const key = x.transaction_date.slice(0, 7)

      if (!result[key]) {
        result[key] = {
          Redemption: 0,
          SWP: 0,
          Switch: 0,
          STP: 0
        }
      }

      const type =
        x.classified_transaction_type

      if (
        result[key] &&
        result[key][type] !== undefined
      ) {
        result[key][type] +=
          Number(x.amount || 0)
      }
    })

    return Object.entries(result)
      .sort((a, b) =>
        a[0].localeCompare(b[0])
      )
      .slice(-8)
  }, [filtered])

  async function uploadFile(e) {
    const file = e.target.files?.[0]

    if (!file) {
      return
    }

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
            defval: null
          }
        )

      const mapped = raw
        .map(mapRow)
        .filter(
          x =>
            x.investor_name &&
            x.transaction_date &&
            x.amount !== null
        )

      if (!mapped.length) {
        throw new Error(
          'No valid transactions found in this Excel file.'
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
        const batch =
          mapped.slice(i, i + 500)

        const { error } =
          await supabase
            .from('transactions')
            .insert(batch)

        if (error) {
          throw error
        }
      }

      setMessage(
        'Analysing transactions and identifying SWP, Redemption, Switch and STP...'
      )

      const { error: rpcError } =
        await supabase.rpc(
          'run_redemption_classification'
        )

      if (rpcError) {
        throw rpcError
      }

      setMessage(
        'Upload completed and transactions classified successfully.'
      )

      await loadTransactions()
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
      Source: x.original_transaction_type,
      Classification:
        x.classified_transaction_type
    }))

    const worksheet =
      XLSX.utils.json_to_sheet(out)

    const workbook =
      XLSX.utils.book_new()

    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'Transaction Data'
    )

    XLSX.writeFile(
      workbook,
      'snowball-transaction-data.xlsx'
    )
  }

  function exportPDF() {
    const doc = new jsPDF({
      orientation: 'landscape'
    })

    doc.setFontSize(16)

    doc.text(
      'Snowball Financial Services - Transaction Data',
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

      body: filtered.map(x => [
        x.transaction_date,
        x.rm_name,
        x.investor_name,
        (x.scheme || '').slice(0, 35),
        money(x.amount),
        x.original_transaction_type || '-',
        x.classified_transaction_type || '-'
      ])
    })

    doc.save(
      'snowball-transaction-data.pdf'
    )
  }

  async function addRM() {
    const name = newRM.trim()

    if (!name) {
      return
    }

    setError('')
    setMessage('')

    const { error } =
      await supabase
        .from('relationship_managers')
        .insert({
          rm_name: name,
          status: 'Active'
        })

    if (error) {
      setError(error.message)
      return
    }

    setNewRM('')
    setMessage('RM added successfully.')

    await loadRMs()
  }

  function startEditRM(item) {
    setEditingRM(item)
    setEditRMName(item.rm_name)
  }

  async function saveRMEdit() {
    if (
      !editingRM ||
      !editRMName.trim()
    ) {
      return
    }

    const { error } =
      await supabase
        .from('relationship_managers')
        .update({
          rm_name:
            editRMName.trim()
        })
        .eq(
          'id',
          editingRM.id
        )

    if (error) {
      setError(error.message)
      return
    }

    setEditingRM(null)
    setEditRMName('')
    setMessage(
      'RM updated successfully.'
    )

    await loadRMs()
  }

  async function toggleRMStatus(item) {
    const newStatus =
      item.status === 'Inactive'
        ? 'Active'
        : 'Inactive'

    const { error } =
      await supabase
        .from('relationship_managers')
        .update({
          status: newStatus
        })
        .eq('id', item.id)

    if (error) {
      setError(error.message)
      return
    }

    setMessage(
      `${item.rm_name} is now ${newStatus}.`
    )

    await loadRMs()
  }

  async function addAdmin() {
    const name =
      newAdminName.trim()

    const adminEmail =
      newAdminEmail.trim().toLowerCase()

    if (!adminEmail) {
      setError(
        'Please enter an admin email address.'
      )
      return
    }

    const { error } =
      await supabase
        .from('admin_users')
        .insert({
          name: name || null,
          email: adminEmail,
          status: 'Active'
        })

    if (error) {
      setError(error.message)
      return
    }

    setNewAdminName('')
    setNewAdminEmail('')

    setMessage(
      'Admin access email added successfully.'
    )

    await loadAdmins()
  }

  async function toggleAdmin(item) {
    const newStatus =
      item.status === 'Inactive'
        ? 'Active'
        : 'Inactive'

    const { error } =
      await supabase
        .from('admin_users')
        .update({
          status: newStatus
        })
        .eq('id', item.id)

    if (error) {
      setError(error.message)
      return
    }

    await loadAdmins()
  }

  function setQuickPeriod(p) {
    setPeriod(p)
    setFrom('')
    setTo('')
  }

  const sourceType = x => {
    const value =
      String(
        x.original_transaction_type ||
        x.transaction_type ||
        ''
      ).toLowerCase()

    if (value.includes('switch')) {
      return 'Switch'
    }

    if (value.includes('stp')) {
      return 'STP'
    }

    if (value.includes('swp')) {
      return 'SWP'
    }

    return 'Redemption'
  }

  if (!session) {
    return (
      <main className="loginPage">

        <section className="loginCard">

          <div className="loginBrand">
            <img
              src="/src/logo.png"
              alt="Snowball Financial Services"
              className="loginLogo"
            />

            <h1>
              Snowball Redemption Tracker
            </h1>

            <p>
              Sign in to monitor redemption activity
            </p>
          </div>

          {!showForgot ? (
            <form
              className="loginForm"
              onSubmit={login}
            >
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
                className="primaryButton fullButton"
                disabled={loading}
              >
                {loading
                  ? 'Signing in...'
                  : 'Login'}
              </button>

              <button
                type="button"
                className="forgotLink"
                onClick={() =>
                  setShowForgot(true)
                }
              >
                Forgot password?
              </button>
            </form>
          ) : (
            <form
              className="loginForm"
              onSubmit={forgotPassword}
            >
              <input
                type="email"
                placeholder="Email address"
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
                {loading
                  ? 'Sending...'
                  : 'Send Reset Link'}
              </button>

              <button
                type="button"
                className="forgotLink"
                onClick={() =>
                  setShowForgot(false)
                }
              >
                Back to Login
              </button>
            </form>
          )}

          {error && (
            <div className="formError">
              {error}
            </div>
          )}

          {message && (
            <div className="formMessage">
              {message}
            </div>
          )}

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
    },
    {
      id: 'settings',
      label: 'Settings',
      icon: Settings
    }
  ]

  return (
    <div className="appShell">

      <aside className="sidebar">

        <div className="brandArea">
          <img
            src="/src/logo.png"
            alt="Snowball Financial Services"
            className="sidebarLogo"
          />
        </div>

        <nav className="sideNav">
          {navItems.map(item => {
            const Icon = item.icon

            return (
              <button
                key={item.id}
                className={
                  activePage === item.id
                    ? 'navItem active'
                    : 'navItem'
                }
                onClick={() =>
                  setActivePage(item.id)
                }
              >
                <Icon size={19} />
                <span>{item.label}</span>
              </button>
            )
          })}
        </nav>

        <div className="sidebarFooter">
          Snowball Financial Services
        </div>

      </aside>

      <main className="mainContent">

        <header className="topHeader">
          <div>
            <h1>
              {activePage === 'dashboard' &&
                'Snowball Redemption Tracker'}

              {activePage === 'transactions' &&
                'Transaction Data'}

              {activePage === 'rms' &&
                'Manage Relationship Managers'}

              {activePage === 'settings' &&
                'Settings'}
            </h1>

            <p>
              {activePage === 'dashboard' &&
                'Analyse transactions and monitor redemption activity'}

              {activePage === 'transactions' &&
                'Upload, review and analyse transaction data'}

              {activePage === 'rms' &&
                'Manage relationship managers in your organisation'}

              {activePage === 'settings' &&
                'Manage administrator access'}
            </p>
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
          <div className="globalError">
            {error}
            <button
              onClick={() =>
                setError('')
              }
            >
              <X size={16} />
            </button>
          </div>
        )}

        {message && (
          <div className="globalMessage">
            {message}
            <button
              onClick={() =>
                setMessage('')
              }
            >
              <X size={16} />
            </button>
          </div>
        )}

        {activePage !== 'rms' &&
          activePage !== 'settings' && (
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
                      ? 'periodButton active'
                      : 'periodButton'
                  }
                  onClick={() =>
                    setQuickPeriod(p)
                  }
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
              <option value="All">
                All RMs
              </option>

              {activeRMNames.map(name => (
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

            {activePage === 'transactions' && (
              <div className="transactionActions">

                <label className="primaryButton uploadButton">
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

              </div>
            )}

            <button
              className="iconButton"
              onClick={loadAllData}
              title="Refresh"
            >
              <RefreshCw size={17} />
            </button>

          </section>
        )}

        {activePage === 'dashboard' && (
          <Dashboard
            totals={totals}
            rmChart={rmChart}
            monthlyData={monthlyData}
            recentRows={filtered.slice(0, 6)}
          />
        )}

        {activePage === 'transactions' && (
          <TransactionPage
            rows={filtered}
            sourceType={sourceType}
          />
        )}

        {activePage === 'rms' && (
          <ManageRMs
            rms={rms}
            newRM={newRM}
            setNewRM={setNewRM}
            addRM={addRM}
            startEditRM={startEditRM}
            toggleRMStatus={toggleRMStatus}
            editingRM={editingRM}
            editRMName={editRMName}
            setEditRMName={setEditRMName}
            saveRMEdit={saveRMEdit}
            cancelEdit={() =>
              setEditingRM(null)
            }
          />
        )}

        {activePage === 'settings' && (
          <SettingsPage
            admins={admins}
            newAdminName={newAdminName}
            setNewAdminName={
              setNewAdminName
            }
            newAdminEmail={
              newAdminEmail
            }
            setNewAdminEmail={
              setNewAdminEmail
            }
            addAdmin={addAdmin}
            toggleAdmin={toggleAdmin}
          />
        )}

      </main>
    </div>
  )
}

function Dashboard({
  totals,
  rmChart,
  monthlyData,
  recentRows
}) {
  const maxRM =
    Math.max(
      ...rmChart.map(x => x[1]),
      1
    )

  const classItems = [
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

  const maxClass =
    Math.max(
      ...classItems.map(x => x.value),
      1
    )

  return (
    <>
      <section className="summaryGrid">

        <SummaryCard
          label="Redemption"
          value={money(totals.Redemption)}
          type="redemption"
        />

        <SummaryCard
          label="SWP"
          value={money(totals.SWP)}
          type="swp"
        />

        <SummaryCard
          label="Switch"
          value={money(totals.Switch)}
          type="switch"
        />

        <SummaryCard
          label="STP"
          value={money(totals.STP)}
          type="stp"
        />

        <SummaryCard
          label="Investors"
          value={totals.Investors}
          type="investors"
        />

        <SummaryCard
          label="Transactions"
          value={totals.Transactions}
          type="transactions"
        />

      </section>

      <section className="dashboardGrid">

        <article className="panel classificationPanel">
          <h2>Amount by Classification</h2>

          <div className="classificationList">
            {classItems.map(item => (
              <div
                key={item.label}
                className="classificationRow"
              >
                <div className="classificationLabel">
                  <span
                    className={`dot ${item.className}`}
                  />
                  {item.label}
                </div>

                <div className="classificationBarWrap">
                  <div
                    className={`classificationBar ${item.className}`}
                    style={{
                      width:
                        `${Math.max(
                          4,
                          (
                            item.value /
                            maxClass
                          ) * 100
                        )}%`
                    }}
                  />
                </div>

                <strong>
                  {money(item.value)}
                </strong>
              </div>
            ))}
          </div>
        </article>

        <article className="panel monthlyPanel">
          <h2>
            Monthly Trend (Amount in ₹)
          </h2>

          <div className="trendChart">
            {monthlyData.map(([month, values]) => {
              const total =
                Math.max(
                  values.Redemption,
                  values.SWP,
                  values.Switch,
                  values.STP,
                  1
                )

              return (
                <div
                  className="monthGroup"
                  key={month}
                >
                  <div className="monthBars">

                    <span
                      className="miniBar redemption"
                      style={{
                        height:
                          `${(
                            values.Redemption /
                            total
                          ) * 80 + 5}px`
                      }}
                    />

                    <span
                      className="miniBar swp"
                      style={{
                        height:
                          `${(
                            values.SWP /
                            total
                          ) * 80 + 5}px`
                      }}
                    />

                    <span
                      className="miniBar switch"
                      style={{
                        height:
                          `${(
                            values.Switch /
                            total
                          ) * 80 + 5}px`
                      }}
                    />

                    <span
                      className="miniBar stp"
                      style={{
                        height:
                          `${(
                            values.STP /
                            total
                          ) * 80 + 5}px`
                      }}
                    />

                  </div>

                  <small>
                    {month.slice(5)}
                  </small>
                </div>
              )
            })}
          </div>
        </article>

      </section>

      <section className="dashboardGrid bottomGrid">

        <article className="panel rmPanel">
          <h2>Transactions by RM</h2>

          <div className="rmBars">
            {rmChart.map(([name, value]) => (
              <div
                className="rmBarRow"
                key={name}
              >
                <span className="rmName">
                  {name}
                </span>

                <div className="rmBarTrack">
                  <div
                    className="rmBar"
                    style={{
                      width:
                        `${(
                          value /
                          maxRM
                        ) * 100}%`
                    }}
                  />
                </div>

                <strong>
                  {money(value)}
                </strong>
              </div>
            ))}
          </div>
        </article>

        <article className="panel recentPanel">
          <div className="panelHeading">
            <h2>
              Recent Transactions
            </h2>

            <span>
              Dashboard overview
            </span>
          </div>

          <div className="recentTableWrap">
            <table>
              <thead>
                <tr>
                  <th>Date</th>
                  <th>RM</th>
                  <th>Investor</th>
                  <th>Amount</th>
                  <th>Classification</th>
                </tr>
              </thead>

              <tbody>
                {recentRows.map(x => (
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
                      {money(x.amount)}
                    </td>

                    <td>
                      <ClassificationBadge
                        value={
                          x.classified_transaction_type
                        }
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </article>

      </section>
    </>
  )
}

function SummaryCard({
  label,
  value,
  type
}) {
  return (
    <article
      className={`summaryCard ${type}`}
    >
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  )
}

function ClassificationBadge({ value }) {
  const type =
    String(value || '')
      .toLowerCase()
      .replace(/\s/g, '')

  return (
    <span
      className={`classificationBadge ${type}`}
    >
      {value || 'Pending'}
    </span>
  )
}

function TransactionPage({
  rows,
  sourceType
}) {
  return (
    <section className="panel transactionPanel">

      <div className="panelHeading">
        <div>
          <h2>
            Transaction Details
          </h2>

          <p>
            {rows.length} transactions
          </p>
        </div>
      </div>

      <div className="dataTableWrap">
        <table className="dataTable">
          <thead>
            <tr>
              <th>Date</th>
              <th>RM</th>
              <th>Investor</th>
              <th>Scheme</th>
              <th>Amount</th>
              <th>Source</th>
              <th>System Classification</th>
            </tr>
          </thead>

          <tbody>
            {rows.map(x => (
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
                  {x.scheme}
                </td>

                <td>
                  {money(x.amount)}
                </td>

                <td>
                  {sourceType(x)}
                </td>

                <td>
                  <ClassificationBadge
                    value={
                      x.classified_transaction_type
                    }
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

    </section>
  )
}

function ManageRMs({
  rms,
  newRM,
  setNewRM,
  addRM,
  startEditRM,
  toggleRMStatus,
  editingRM,
  editRMName,
  setEditRMName,
  saveRMEdit,
  cancelEdit
}) {
  return (
    <div className="managePage">

      <section className="panel addRMPanel">
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
            className="primaryButton"
            onClick={addRM}
          >
            <Plus size={17} />
            Add RM
          </button>
        </div>

        <p>
          Inactive RMs remain in historical transaction data.
        </p>
      </section>

      <section className="panel currentRMPanel">
        <h2>Current RMs</h2>

        <div className="rmList">

          {rms.map(item => (
            <div
              className="rmListItem"
              key={item.id}
            >

              {editingRM?.id === item.id ? (
                <>
                  <input
                    value={editRMName}
                    onChange={e =>
                      setEditRMName(
                        e.target.value
                      )
                    }
                  />

                  <div className="rmActions">
                    <button
                      className="smallPrimary"
                      onClick={saveRMEdit}
                    >
                      Save
                    </button>

                    <button
                      className="smallSecondary"
                      onClick={cancelEdit}
                    >
                      Cancel
                    </button>
                  </div>
                </>
              ) : (
                <>
                  <div className="rmInfo">
                    <strong>
                      {item.rm_name}
                    </strong>

                    <span
                      className={
                        item.status === 'Inactive'
                          ? 'status inactive'
                          : 'status active'
                      }
                    >
                      {item.status || 'Active'}
                    </span>
                  </div>

                  <div className="rmActions">

                    <button
                      className="smallSecondary"
                      onClick={() =>
                        startEditRM(item)
                      }
                    >
                      <Pencil size={15} />
                      Edit
                    </button>

                    <button
                      className={
                        item.status === 'Inactive'
                          ? 'activateButton'
                          : 'inactiveButton'
                      }
                      onClick={() =>
                        toggleRMStatus(item)
                      }
                    >
                      {item.status === 'Inactive'
                        ? (
                          <>
                            <UserCheck size={15} />
                            Activate
                          </>
                        )
                        : (
                          <>
                            <UserX size={15} />
                            Inactive
                          </>
                        )}
                    </button>

                  </div>
                </>
              )}

            </div>
          ))}

        </div>
      </section>

    </div>
  )
}

function SettingsPage({
  admins,
  newAdminName,
  setNewAdminName,
  newAdminEmail,
  setNewAdminEmail,
  addAdmin,
  toggleAdmin
}) {
  return (
    <div className="settingsPage">

      <section className="panel adminPanel">

        <div className="panelHeading">
          <div>
            <h2>
              Admin Access
            </h2>

            <p>
              Manage email IDs authorised to access the administration area.
            </p>
          </div>

          <ShieldCheck size={28} />
        </div>

        <div className="adminAddGrid">

          <input
            placeholder="Admin name (optional)"
            value={newAdminName}
            onChange={e =>
              setNewAdminName(
                e.target.value
              )
            }
          />

          <input
            type="email"
            placeholder="Admin email address"
            value={newAdminEmail}
            onChange={e =>
              setNewAdminEmail(
                e.target.value
              )
            }
          />

          <button
            className="primaryButton"
            onClick={addAdmin}
          >
            <Plus size={17} />
            Add Admin
          </button>

        </div>

        <div className="adminList">

          {admins.map(item => (
            <div
              className="adminListItem"
              key={item.id}
            >
              <div>
                <strong>
                  {item.name || 'Administrator'}
                </strong>

                <span>
                  {item.email}
                </span>
              </div>

              <button
                className={
                  item.status === 'Inactive'
                    ? 'activateButton'
                    : 'inactiveButton'
                }
                onClick={() =>
                  toggleAdmin(item)
                }
              >
                {item.status === 'Inactive'
                  ? 'Activate'
                  : 'Inactive'}
              </button>
            </div>
          ))}

        </div>

      </section>
    </div>
  )
}

createRoot(
  document.getElementById('root')
).render(<App />)
