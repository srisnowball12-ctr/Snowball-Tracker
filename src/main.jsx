import React, { useEffect, useMemo, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { createClient } from '@supabase/supabase-js'
import * as XLSX from 'xlsx'
import jsPDF from 'jspdf'
import autoTable from 'jspdf-autotable'
import { Upload, Download, FileText, LogOut, RefreshCw } from 'lucide-react'
import './styles.css'

const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
)

const money = n => new Intl.NumberFormat('en-IN',{style:'currency',currency:'INR',maximumFractionDigits:0}).format(Number(n||0))
const iso = v => {
  if (!v) return null
  if (v instanceof Date) return v.toISOString().slice(0,10)
  const d = new Date(v); return isNaN(d) ? null : d.toISOString().slice(0,10)
}
const norm = s => String(s||'').trim().toLowerCase().replace(/[^a-z0-9]/g,'')

function mapRow(row) {
  const lookup = Object.fromEntries(Object.entries(row).map(([k,v]) => [norm(k),v]))
  const get = (...keys) => keys.map(k=>lookup[norm(k)]).find(v => v !== undefined)
  const amountRaw = get('Amount(₹)','Amount','amount')
  const amount = typeof amountRaw === 'number' ? amountRaw : Number(String(amountRaw||'').replace(/[₹,\s]/g,''))
  return {
    rm_name: get('Partner/Employee','RM','rm_name') || null,
    group_name: get('Group','group_name') || null,
    investor_name: get('Investor','investor_name') || null,
    transaction_date: iso(get('Date','transaction_date')),
    folio_no: String(get('Folio No/Demat A/C','Folio','folio_no') || '') || null,
    scheme: get('Scheme','Fund','scheme') || null,
    amount: Number.isFinite(amount) ? amount : null,
    transaction_type: 'Imported',
    original_transaction_type: get('Type','original_transaction_type') || null,
    classification_status: 'Needs Review',
    classification_reason: get('Reason','classification_reason') || null
  }
}

function App(){
  const [session,setSession]=useState(null)
  const [email,setEmail]=useState('')
  const [password,setPassword]=useState('')
  const [error,setError]=useState('')
  const [loading,setLoading]=useState(false)
  const [rows,setRows]=useState([])
  const [rms,setRms]=useState([])
  const [rm,setRm]=useState('All')
  const [from,setFrom]=useState('')
  const [to,setTo]=useState('')
  const [period,setPeriod]=useState('YTD')
  const [uploading,setUploading]=useState(false)
  const [message,setMessage]=useState('')

  useEffect(()=>{ supabase.auth.getSession().then(({data})=>setSession(data.session)); const {data:{subscription}}=supabase.auth.onAuthStateChange((_e,s)=>setSession(s)); return ()=>subscription.unsubscribe()},[])
  useEffect(()=>{ if(session) loadData() },[session])

  async function login(e){
    e.preventDefault(); setLoading(true); setError('')
    const {error} = await supabase.auth.signInWithPassword({email,password})
    if(error) setError(error.message)
    setLoading(false)
  }

  async function loadData(){
    setLoading(true)
    const {data,error} = await supabase.from('transactions').select('*').order('transaction_date',{ascending:false})
    if(error) setError(error.message); else { setRows(data||[]); setRms(['All',...Array.from(new Set((data||[]).map(x=>x.rm_name).filter(Boolean))).sort()]) }
    setLoading(false)
  }

  const filtered = useMemo(()=>rows.filter(x=>{
    if(rm!=='All' && x.rm_name!==rm) return false
    const d=x.transaction_date
    if(from && d<from) return false
    if(to && d>to) return false
    if(!from && !to){
      const now=new Date(); const dt=new Date(d+'T00:00:00')
      if(period==='WTD'){ const day=(now.getDay()+6)%7; const start=new Date(now); start.setDate(now.getDate()-day); start.setHours(0,0,0,0); if(dt<start)return false }
      if(period==='MTD' && (dt.getMonth()!==now.getMonth()||dt.getFullYear()!==now.getFullYear()))return false
      if(period==='QTD'){ const q=Math.floor(now.getMonth()/3); if(dt.getFullYear()!==now.getFullYear()||Math.floor(dt.getMonth()/3)!==q)return false }
      if(period==='YTD' && dt.getFullYear()!==now.getFullYear())return false
    }
    return true
  }),[rows,rm,from,to,period])

  const totals = useMemo(()=>({
    Redemption: filtered.filter(x=>x.classified_transaction_type==='Redemption').reduce((s,x)=>s+Number(x.amount||0),0),
    SWP: filtered.filter(x=>x.classified_transaction_type==='SWP').reduce((s,x)=>s+Number(x.amount||0),0),
    Switch: filtered.filter(x=>x.classified_transaction_type==='Switch').reduce((s,x)=>s+Number(x.amount||0),0),
    STP: filtered.filter(x=>x.classified_transaction_type==='STP').reduce((s,x)=>s+Number(x.amount||0),0),
    Investors: new Set(filtered.map(x=>x.investor_name).filter(Boolean)).size,
    Transactions: filtered.length
  }),[filtered])

  async function uploadFile(e){
    const file=e.target.files?.[0]; if(!file)return
    setUploading(true); setMessage('Reading Excel...')
    try{
      const buf=await file.arrayBuffer()
      const wb=XLSX.read(buf,{type:'array',cellDates:true})
      const ws=wb.Sheets[wb.SheetNames[0]]
      const raw=XLSX.utils.sheet_to_json(ws,{defval:null})
      const mapped=raw.map(mapRow).filter(r=>r.investor_name&&r.transaction_date&&r.amount!=null)
      if(!mapped.length) throw new Error('No valid transactions found. Please use the normal Snowball redemption Excel format.')
      setMessage(`Uploading ${mapped.length} transactions...`)
      for(let i=0;i<mapped.length;i+=500){
        const {error}=await supabase.from('transactions').insert(mapped.slice(i,i+500))
        if(error) throw error
      }
      setMessage('Analysing SWP vs Redemption patterns...')
      const {error:rpcError}=await supabase.rpc('run_redemption_classification')
      if(rpcError) throw rpcError
      setMessage('Done. Dashboard updated.')
      await loadData()
    }catch(err){ setError(err.message); setMessage('') }
    setUploading(false); e.target.value=''
  }

  function exportExcel(){
    const out=filtered.map(x=>({Date:x.transaction_date,RM:x.rm_name,Investor:x.investor_name,Folio:x.folio_no,Scheme:x.scheme,Amount:x.amount,SourceType:x.original_transaction_type,FinalClassification:x.classified_transaction_type,Reason:x.classification_reason}))
    const ws=XLSX.utils.json_to_sheet(out); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'Redemption Tracker'); XLSX.writeFile(wb,'snowball-redemption-report.xlsx')
  }
  function exportPDF(){
    const doc=new jsPDF({orientation:'landscape'})
    doc.setFontSize(16); doc.text('Snowball Redemption Tracker',14,14)
    doc.setFontSize(10); doc.text(`RM: ${rm} | Period: ${from||to?`${from||''} to ${to||''}`:period}`,14,21)
    autoTable(doc,{startY:27,head:[['Date','RM','Investor','Scheme','Amount','Classification']],body:filtered.map(x=>[x.transaction_date,x.rm_name,x.investor_name,(x.scheme||'').slice(0,35),money(x.amount),x.classified_transaction_type])})
    doc.save('snowball-redemption-report.pdf')
  }

if(!session) return <main className="login"><section><h1>Snowball Financial Services</h1><p>Redemption Tracker</p><form onSubmit={login}><input placeholder="Email" value={email} onChange={e=>setEmail(e.target.value)} /><input type="password" placeholder="Password" value={password} onChange={e=>setPassword(e.target.value)} /><button disabled={loading}>{loading?'Signing in...':'Login'}</button>{error&&<small>{error}</small>}</form></section></main>
  
  return <main>
    <header><div><h1>Snowball Redemption Tracker</h1><p>Analyse redemptions • Identify SWPs • Track RM-wise activity</p></div><button className="ghost" onClick={()=>supabase.auth.signOut()}><LogOut size={17}/> Logout</button></header>
    {error&&<div className="error">{error}</div>}
    <section className="toolbar">
      <div className="periods">{['WTD','MTD','QTD','YTD'].map(p=><button className={period===p&&!from&&!to?'active':''} onClick={()=>{setPeriod(p);setFrom('');setTo('')}} key={p}>{p}</button>)}</div>
      <select value={rm} onChange={e=>setRm(e.target.value)}>{rms.map(x=><option key={x}>{x}</option>)}</select>
      <input type="date" value={from} onChange={e=>setFrom(e.target.value)} />
      <input type="date" value={to} onChange={e=>setTo(e.target.value)} />
      <button className="upload"><Upload size={17}/> <label>Upload Excel<input type="file" accept=".xlsx,.xls,.csv" onChange={uploadFile}/></label></button>
      <button onClick={exportExcel}><Download size={17}/> Excel</button><button onClick={exportPDF}><FileText size={17}/> PDF</button><button onClick={loadData}><RefreshCw size={17}/></button>
    </section>
    {message&&<div className="message">{message}</div>}
    <section className="cards">
      {Object.entries(totals).map(([k,v])=><article key={k}><span>{k}</span><strong>{['Investors','Transactions'].includes(k)?v:money(v)}</strong></article>)}
    </section>
    <section className="tableWrap"><h2>Transactions ({filtered.length})</h2><table><thead><tr><th>Date</th><th>RM</th><th>Investor</th><th>Scheme</th><th>Amount</th><th>Source</th><th>System Classification</th></tr></thead><tbody>{filtered.slice(0,500).map(x=><tr key={x.id}><td>{x.transaction_date}</td><td>{x.rm_name}</td><td>{x.investor_name}</td><td>{x.scheme}</td><td>{money(x.amount)}</td><td>{x.original_transaction_type}</td><td><b>{x.classified_transaction_type}</b></td></tr>)}</tbody></table></section>
  </main>
}
createRoot(document.getElementById('root')).render(<App/>)
