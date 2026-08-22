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
  Table2
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

const norm = s =>
  String(s || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '')

const iso = v => {
  if (!v) return null

  if (v instanceof Date && !isNaN(v)) {
    return v.toISOString().slice(0, 10)
  }

  const d = new Date(v)

  if (!isNaN(d)) {
    return d.toISOString().slice(0, 10)
  }

  return null
}

/* -------------------------------------------------------
   CLEAN DISPLAY OF ORIGINAL SOURCE TYPE
------------------------------------------------------- */

function sourceLabel(value) {
  const s = String(value || '').trim().toLowerCase()

  if (s.includes('swp')) return 'SWP'
  if (s.includes('switch')) return 'Switch'
  if (s.includes('stp')) return 'STP'

  if (
    s === 'red' ||
    s.includes('redeem') ||
    s.includes('redemption')
  ) {
    return 'Redemption'
  }

  return value || 'Unknown'
}

/* -------------------------------------------------------
   MAP EXCEL ROW
------------------------------------------------------- */

function mapRow(row) {
  const lookup = Object.fromEntries(
    Object.entries(row).map(([k, v]) => [norm(k), v])
  )

  const get = (...keys) =>
    keys
      .map(k => lookup[norm(k)])
      .find(v => v !== undefined && v !== null && v !== '')

  const amountRaw = get('Amount(₹)', 'Amount', 'amount')

  const amount =
    typeof amountRaw === 'number'
      ? amountRaw
      : Number(String(amountRaw || '').replace(/[₹,\s]/g, ''))

  return {
    rm_name: get('Partner/Employee', 'RM', 'rm_name') || null,

    group_name: get('Group', 'group_name') || null,

    investor_name:
      get('Investor', 'Investor Name', 'investor_name') || null,

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

    transaction_type: 'Imported',

    original_transaction_type:
      get('Type', 'Transaction Type', 'original_transaction_type') || null,

    classified_transaction_type: null,

    classification_status: 'Needs Review',

    classification_reason: null
  }
}

/* -------------------------------------------------------
   DATE DIFFERENCE
------------------------------------------------------- */

function daysBetween(a, b) {
  const d1 = new Date(a + 'T00:00:00')
  const d2 = new Date(b + 'T00:00:00')

  return Math.round(
    Math.abs(d2.getTime() - d1.getTime()) /
      (1000 * 60 * 60 * 24)
  )
}

/* -------------------------------------------------------
   MAIN ANALYTICAL ENGINE

   PRIORITY:

   1. Explicit SWP in Excel = SWP
   2. Explicit Switch = Switch
   3. Explicit STP = STP
   4. Generic RED / Redemption is analysed:
      recurring withdrawals from same investor +
      folio + scheme at regular intervals
      are identified as likely SWP.
------------------------------------------------------- */

function analyseTransactions(records) {
  const groups = {}

  records.forEach((row, index) => {
    const key = [
      norm(row.investor_name),
      norm(row.folio_no),
      norm(row.scheme)
    ].join('|')

    if (!groups[key]) groups[key] = []

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
      let reason = 'One-time or non-recurring redemption'

      /* Explicit classifications */

      if (raw.includes('swp')) {
        classification = 'SWP'
        reason = 'SWP explicitly indicated in source data'
      } else if (raw.includes('switch')) {
        classification = 'Switch'
        reason = 'Switch explicitly indicated in source data'
      } else if (raw.includes('stp')) {
        classification = 'STP'
        reason = 'STP explicitly indicated in source data'
      } else {
        /*
          Analytical SWP detection.

          Look for recurring transactions in the same:
          Investor + Folio + Scheme
        */

        const nearbyIntervals = []

        if (position > 0) {
          nearbyIntervals.push(
            daysBetween(
              row.transaction_date,
              group[position - 1].transaction_date
            )
          )
        }

        if (position < group.length - 1) {
          nearbyIntervals.push(
            daysBetween(
              group[position + 1].transaction_date,
              row.transaction_date
            )
          )
        }

        const monthlyPattern =
          nearbyIntervals.filter(
            d => d >= 20 && d <= 40
          ).length

        const quarterlyPattern =
          nearbyIntervals.filter(
            d => d >= 75 && d <= 105
          ).length

        /*
          Strong recurring pattern:
          3 or more transactions in same folio/scheme
          with approximately monthly or quarterly frequency.
        */

        if (
          group.length >= 3 &&
          (monthlyPattern >= 1 || quarterlyPattern >= 1)
        ) {
          classification = 'SWP'

          reason =
            monthlyPattern >= 1
              ? 'System detected recurring approximately monthly withdrawals'
              : 'System detected recurring approximately quarterly withdrawals'
        }

        /*
          Two similar transactions approximately one month apart.
        */

        else if (
          group.length === 2 &&
          monthlyPattern >= 1
        ) {
          const other =
            position === 0 ? group[1] : group[0]

          const currentAmount = Math.abs(
            Number(row.amount || 0)
          )

          const otherAmount = Math.abs(
            Number(other.amount || 0)
          )

          const difference =
            Math.abs(currentAmount - otherAmount)

          const average =
            (currentAmount + otherAmount) / 2

          /*
            If amounts are reasonably similar,
            it is more likely to be an SWP.
          */

          if (
            average > 0 &&
            difference / average <= 0.5
          ) {
            classification = 'SWP'

            reason =
              'System detected recurring monthly withdrawals with similar amounts'
          }
        }
      }

      row.classified_transaction_type =
        classification

      row.classification_status =
        'Auto Classified'

      row.classification_reason =
        reason
    })
  })

  const output = new Array(records.length)

  Object.values(groups).forEach(group => {
    group.forEach(row => {
      const { __index, ...cleanRow } = row
      output[__index] = cleanRow
    })
  })

  return output
}

/* -------------------------------------------------------
   APP
------------------------------------------------------- */

function App() {
  const [session, setSession] = useState(null)

  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')

  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)

  const [rows, setRows] = useState([])
  const [rms, setRms] = useState([])

  const [rm, setRm] = useState('All')

  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')

  const [period, setPeriod] = useState('YTD')

  const [uploading, setUploading] = useState(false)

  const [message, setMessage] = useState('')

  const [activeTab, setActiveTab] =
    useState('dashboard')

  /* Authentication */

  useEffect(() => {
    supabase.auth
      .getSession()
      .then(({ data }) => {
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
    if (session) loadData()
  }, [session])

  /* Login */

  async function login(e) {
    e.preventDefault()

    setLoading(true)
    setError('')

    const { error } =
      await supabase.auth.signInWithPassword({
        email,
        password
      })

    if (error) setError(error.message)

    setLoading(false)
  }

  /* Load data */

  async function loadData() {
    setLoading(true)
    setError('')

    const { data, error } = await supabase
      .from('transactions')
      .select('*')
      .order('transaction_date', {
        ascending: false
      })

    if (error) {
      setError(error.message)
    } else {
      const allRows = data || []

      setRows(allRows)

      setRms([
        'All',
        ...Array.from(
          new Set(
            allRows
              .map(x => x.rm_name)
              .filter(Boolean)
          )
        ).sort()
      ])
    }

    setLoading(false)
  }

  /* -------------------------------------------------------
     FILTER DATA
  ------------------------------------------------------- */

  const filtered = useMemo(() => {
    return rows.filter(x => {
      if (
        rm !== 'All' &&
        x.rm_name !== rm
      ) {
        return false
      }

      const d = x.transaction_date

      if (from && d < from) return false
      if (to && d > to) return false

      if (!from && !to && d) {
        const now = new Date()

        const dt = new Date(
          d + 'T00:00:00'
        )

        if (period === 'WTD') {
          const day =
            (now.getDay() + 6) % 7

          const start = new Date(now)

          start.setDate(
            now.getDate() - day
          )

          start.setHours(
            0,
            0,
            0,
            0
          )

          if (dt < start) return false
        }

        if (
          period === 'MTD' &&
          (
            dt.getMonth() !==
              now.getMonth() ||
            dt.getFullYear() !==
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
              dt.getMonth() / 3
            )

          if (
            dt.getFullYear() !==
              now.getFullYear() ||
            transactionQuarter !==
              currentQuarter
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

  /* Dashboard totals */

  const totals = useMemo(() => {
    return {
      Redemption: filtered
        .filter(
          x =>
            x.classified_transaction_type ===
            'Redemption'
        )
        .reduce(
          (sum, x) =>
            sum + Number(x.amount || 0),
          0
        ),

      SWP: filtered
        .filter(
          x =>
            x.classified_transaction_type ===
            'SWP'
        )
        .reduce(
          (sum, x) =>
            sum + Number(x.amount || 0),
          0
        ),

      Switch: filtered
        .filter(
          x =>
            x.classified_transaction_type ===
            'Switch'
        )
        .reduce(
          (sum, x) =>
            sum + Number(x.amount || 0),
          0
        ),

      STP: filtered
        .filter(
          x =>
            x.classified_transaction_type ===
            'STP'
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

  /* -------------------------------------------------------
     UPLOAD EXCEL
  ------------------------------------------------------- */

  async function uploadFile(e) {
    const file = e.target.files?.[0]

    if (!file) return

    setUploading(true)
    setError('')
    setMessage('Reading Excel...')

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
          r =>
            r.investor_name &&
            r.transaction_date &&
            r.amount !== null
        )

      if (!mapped.length) {
        throw new Error(
          'No valid transactions found. Please use the normal Snowball redemption Excel format.'
        )
      }

      /*
        Analyse existing history + new Excel.

        This allows the system to identify recurring
        transactions even when the Excel says only RED.
      */

      setMessage(
        'Analysing transaction patterns...'
      )

      const combined =
        [...rows, ...mapped]

      const analysed =
        analyseTransactions(combined)

      const analysedNewRows =
        analysed.slice(rows.length)

      setMessage(
        `Uploading ${analysedNewRows.length} analysed transactions...`
      )

      /*
        Insert in batches of 500
      */

      for (
        let i = 0;
        i < analysedNewRows.length;
        i += 500
      ) {
        const batch =
          analysedNewRows.slice(
            i,
            i + 500
          )

        const { error } =
          await supabase
            .from('transactions')
            .insert(batch)

        if (error) throw error
      }

      setMessage(
        'Done. Dashboard updated.'
      )

      await loadData()
      setActiveTab('dashboard')
    } catch (err) {
      setError(err.message)
      setMessage('')
    }

    setUploading(false)
    e.target.value = ''
  }

  /* -------------------------------------------------------
     RE-ANALYSE EXISTING DATA

     This is important because your current 1,000
     transactions were classified using the earlier logic.
  ------------------------------------------------------- */

  async function reanalyseAll() {
    if (!rows.length) return

    setUploading(true)
    setError('')
    setMessage(
      'Re-analysing all existing transactions...'
    )

    try {
      const analysed =
        analyseTransactions(rows)

      let updated = 0

      for (
        let i = 0;
        i < analysed.length;
        i += 100
      ) {
        const batch =
          analysed.slice(i, i + 100)

        await Promise.all(
          batch.map(async row => {
            const { error } =
              await supabase
                .from('transactions')
                .update({
                  classified_transaction_type:
                    row.classified_transaction_type,

                  classification_status:
                    row.classification_status,

                  classification_reason:
                    row.classification_reason
                })
                .eq('id', row.id)

            if (error) throw error
          })
        )

        updated += batch.length

        setMessage(
          `Re-analysing transactions... ${updated} of ${analysed.length}`
        )
      }

      setMessage(
        'Re-analysis complete. Dashboard updated.'
      )

      await loadData()
    } catch (err) {
      setError(err.message)
      setMessage('')
    }

    setUploading(false)
  }

  /* -------------------------------------------------------
     EXPORT EXCEL
  ------------------------------------------------------- */

  function exportExcel() {
    const output = filtered.map(x => ({
      Date: x.transaction_date,
      RM: x.rm_name,
      Investor: x.investor_name,
      Folio: x.folio_no,
      Scheme: x.scheme,
      Amount: x.amount,
      Source: sourceLabel(
        x.original_transaction_type
      ),
      SystemClassification:
        x.classified_transaction_type,
      ClassificationReason:
        x.classification_reason
    }))

    const worksheet =
      XLSX.utils.json_to_sheet(output)

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

  /* -------------------------------------------------------
     EXPORT PDF
  ------------------------------------------------------- */

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

      body: filtered.map(x => [
        x.transaction_date,
        x.rm_name,
        x.investor_name,
        (x.scheme || '').slice(0, 35),
        money(x.amount),
        sourceLabel(
          x.original_transaction_type
        ),
        x.classified_transaction_type
      ])
    })

    doc.save(
      'snowball-redemption-report.pdf'
    )
  }

  /* -------------------------------------------------------
     LOGIN SCREEN
  ------------------------------------------------------- */

  if (!session) {
    return (
      <main className="login">
        <section>
          <h1>
            Snowball Financial Services
          </h1>

          <p>Redemption Tracker</p>

          <form onSubmit={login}>
            <input
              placeholder="Email"
              value={email}
              onChange={e =>
                setEmail(e.target.value)
              }
            />

            <input
              type="password"
              placeholder="Password"
              value={password}
              onChange={e =>
                setPassword(e.target.value)
              }
            />

            <button disabled={loading}>
              {loading
                ? 'Signing in...'
                : 'Login'}
            </button>

            {error && (
              <small>{error}</small>
            )}
          </form>
        </section>
      </main>
    )
  }

  /* -------------------------------------------------------
     MAIN APPLICATION
  ------------------------------------------------------- */

  return (
    <main>
      <header>
        <div>
          <h1>
            Snowball Redemption Tracker
          </h1>

          <p>
            Analyse redemptions • Identify SWPs •
            Track RM-wise activity
          </p>
        </div>

        <button
          className="ghost"
          onClick={() =>
            supabase.auth.signOut()
          }
        >
          <LogOut size={17} />
          Logout
        </button>
      </header>

      {error && (
        <div className="error">
          {error}
        </div>
      )}

      {/* MAIN TABS */}

      <section className="toolbar">
        <button
          className={
            activeTab === 'dashboard'
              ? 'active'
              : ''
          }
          onClick={() =>
            setActiveTab('dashboard')
          }
        >
          <LayoutDashboard size={17} />
          Dashboard
        </button>

        <button
          className={
            activeTab === 'transactions'
              ? 'active'
              : ''
          }
          onClick={() =>
            setActiveTab('transactions')
          }
        >
          <Table2 size={17} />
          Transaction Data
        </button>
      </section>

      {/* FILTERS AND ACTIONS */}

      <section className="toolbar">
        <div className="periods">
          {[
            'WTD',
            'MTD',
            'QTD',
            'YTD'
          ].map(p => (
            <button
              className={
                period === p &&
                !from &&
                !to
                  ? 'active'
                  : ''
              }

              onClick={() => {
                setPeriod(p)
                setFrom('')
                setTo('')
              }}

              key={p}
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
          {rms.map(x => (
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

        <button className="upload">
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

        <button
          onClick={reanalyseAll}
          disabled={uploading}
          title="Re-analyse all transactions"
        >
          <RefreshCw size={17} />
        </button>
      </section>

      {message && (
        <div className="message">
          {message}
        </div>
      )}

      {/* DASHBOARD */}

      {activeTab === 'dashboard' && (
        <>
          <section className="cards">
            {Object.entries(totals).map(
              ([key, value]) => (
                <article key={key}>
                  <span>{key}</span>

                  <strong>
                    {[
                      'Investors',
                      'Transactions'
                    ].includes(key)
                      ? value
                      : money(value)}
                  </strong>
                </article>
              )
            )}
          </section>
        </>
      )}

      {/* TRANSACTION DATA */}

      {activeTab === 'transactions' && (
        <section className="tableWrap">
          <h2>
            Transaction Data (
            {filtered.length})
          </h2>

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
                <th>
                  System Classification
                </th>
                <th>
                  Classification Reason
                </th>
              </tr>
            </thead>

            <tbody>
              {filtered
                .slice(0, 1000)
                .map(x => (
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
                      {sourceLabel(
                        x.original_transaction_type
                      )}
                    </td>

                    <td>
                      <b>
                        {
                          x.classified_transaction_type
                        }
                      </b>
                    </td>

                    <td>
                      {x.classification_reason}
                    </td>
                  </tr>
                ))}
            </tbody>
          </table>
        </section>
      )}
    </main>
  )
}

createRoot(
  document.getElementById('root')
).render(<App />)
