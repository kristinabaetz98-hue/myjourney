
/* MyJourney: Ein-Klick-Zuordnung. Nach outlook.js und cloud.js laden. */
(() => {
  const CONFIG = window.MYJOURNEY_CONFIG || {};
  const URL = 'https://appdqolpwvjktltvtqvj.supabase.co/functions/v1/outlook-import';
  if (!window.supabase?.createClient || !CONFIG.url || !CONFIG.publishableKey) return;
  const client = window.supabase.createClient(CONFIG.url, CONFIG.publishableKey, {auth:{persistSession:true,autoRefreshToken:true,detectSessionInUrl:true}});
  let busy=false, lastPanel=null;
  const el=(tag,text)=>{const x=document.createElement(tag);if(text!==undefined)x.textContent=text;return x;};
  const btn=(text,fn)=>{const x=el('button',text);x.type='button';x.className='secondary';x.style.margin='4px';x.addEventListener('click',fn);return x;};
  async function api(action,extra={}) {
    const {data,error}=await client.auth.getSession();
    if(error || !data.session?.access_token) throw Error('Bitte bei MyJourney anmelden.');
    const r=await fetch(URL,{method:'POST',headers:{'Content-Type':'application/json',apikey:CONFIG.publishableKey,Authorization:'Bearer '+data.session.access_token},body:JSON.stringify({action,...extra})});
    const j=await r.json().catch(()=>({}));if(!r.ok)throw Error(j.error||'Serverfehler');return j;
  }
  function tripMatch(d) {
    const trips=state.trips||[];
    const destination=(d.destination||'').toLocaleLowerCase('de');
    const date=d.start?.slice(0,10)||'';
    const scored=trips.map(t=>{
      const words=[t.destination,t.name].join(' ').toLocaleLowerCase('de');
      const place=destination && words.includes(destination);
      const overlap=date && t.start && t.end && date>=t.start.slice(0,10) && date<=t.end.slice(0,10);
      return {t,score:(place?3:0)+(overlap?2:0)};
    }).sort((a,b)=>b.score-a.score);
    return scored[0]?.score>=2?scored[0].t:null;
  }
  function makeItem(c,tripId){
    const d=c.details||{};
    const start=d.start||'';
    if(d.booking_status==='cancelled')throw Error('Diese Buchung wurde storniert und kann nicht als aktive Reisebuchung übernommen werden.');
    if(!start)throw Error('Es fehlt ein Reisedatum. Bitte diese Buchung später manuell ergänzen.');
    const allowed=['flight','hotel','transfer','activity','other'];
    const type=c.booking_type==='transport'?'transfer':c.booking_type;
    return {id:crypto.randomUUID(),tripId,type:allowed.includes(type)?type:'other',title:d.title||c.subject||'Outlook-Buchung',start,end:d.end||'',location:d.destination||'',address:'',lat:null,lng:null,confirmation:d.confirmation||'',cost:typeof d.cost_eur==='number'?d.cost_eur:0,notes:[`Outlook: ${c.sender||''}`,d.currency && d.amount!=null?`Originalpreis: ${d.amount} ${d.currency}`:'',d.cost_eur==null?'Preis nicht eindeutig erkannt – bitte prüfen.':'',c.preview||''].filter(Boolean).join('\n'),source:'outlook',outlookMessageId:c.microsoft_message_id};
  }
  async function assign(c,trip){
    if(state.items.some(i=>i.outlookMessageId===c.microsoft_message_id)){alert('Diese Buchung ist bereits übernommen.');return;}
    let item;try{item=makeItem(c,trip.id);}catch(e){alert(e.message);return;}
    state.items.push(item);save();render();
    await api('review',{id:c.id,status:'accepted'});
    alert('Buchung zu „'+trip.name+'“ hinzugefügt.');await show();
  }
  async function createTrip(c){
    const d=c.details||{};if(d.booking_status==='cancelled'){alert('Stornierte Buchungen können nicht als aktive Buchung übernommen werden.');return;}if(!d.start){alert('Kein eindeutiges Reisedatum erkannt. Bitte zuerst manuell prüfen.');return;}
    const name=prompt('Wie soll die neue Reise heißen?',d.destination||'Neue Reise');if(!name?.trim())return;
    const start=d.start.slice(0,10),end=(d.end||d.start).slice(0,10);
    const trip={id:crypto.randomUUID(),name:name.trim(),destination:d.destination||name.trim(),start,end:end>=start?end:start};
    let item;try{item=makeItem(c,trip.id);}catch(e){alert(e.message);return;}
    state.trips.push(trip);state.items.push(item);selected=trip.id;save();render();
    await api('review',{id:c.id,status:'accepted'});
    alert('Reise „'+trip.name+'“ angelegt und Buchung übernommen.');await show();
  }
  function addSelect(c,wrap){
    const sel=el('select');sel.style.maxWidth='100%';
    for(const t of state.trips){const opt=el('option',t.name);opt.value=t.id;sel.append(opt);}
    wrap.append(sel,btn('Dieser Reise zuordnen',()=>{const t=state.trips.find(x=>x.id===sel.value);if(t)void assign(c,t).catch(err=>alert(err.message));}));
  }
  async function show(){
    const target=document.querySelector('#outlookSmartResults');if(!target)return;
    target.replaceChildren(el('p','Buchungsvorschläge werden geladen …'));
    try{
      const data=await api('list');target.replaceChildren();
      const candidates=(data.items||[]).filter(c=>c.status==='new');
      target.append(el('h4',`Noch zu prüfen: ${candidates.length} (von ${data.total||0} Treffern)`));
      if(!candidates.length){target.append(el('p','Keine offenen Vorschläge in den zuletzt geladenen Treffern.'));return;}
      const active=candidates.filter(c=>c.details?.booking_status!=='cancelled');
      const cancelled=candidates.filter(c=>c.details?.booking_status==='cancelled');
      if(cancelled.length)target.append(el('p',`Stornierungen separat: ${cancelled.length} – nicht als aktive Buchung übernehmen.`));
      for(const c of [...active,...cancelled]){
        const d=c.details||{};
        const isCancelled=d.booking_status==='cancelled';
        const card=el('article');card.className='card';card.style.margin='10px 0';card.style.padding='16px';
        card.style.borderLeft=isCancelled?'4px solid #b76a42':'4px solid #739d8c';
        card.append(el('strong',`${c.booking_type==='flight'?'✈️':c.booking_type==='hotel'?'🏨':'🎫'} ${d.title||c.subject||'Unbekannte Buchung'}`));
        const status=d.booking_status==='cancelled'?'🔴 Storniert':d.booking_status==='changed'?'🟠 Geändert':'⚪ Status noch nicht sicher erkannt';
        card.append(el('p',status));
        const details=el('div');details.style.display='grid';details.style.gridTemplateColumns='repeat(auto-fit,minmax(180px,1fr))';details.style.gap='8px';
        const field=(name,value)=>{const x=el('div');x.append(el('small',name));x.append(el('div',value||'Nicht erkannt'));return x;};
        details.append(field('📍 Ort / Ziel',d.destination),field('📅 Beginn',d.start),field('📅 Ende',d.end),field('🔖 Buchungsnummer',d.confirmation));
        const cost=d.amount!=null?`${d.amount} ${d.currency||''}`:'Nicht erkannt';
        details.append(field('💳 Gesamtpreis',cost),field('✉️ Absender',c.sender),field('📨 Eingegangen',c.received_at?new Date(c.received_at).toLocaleDateString('de-DE'):'Unbekannt'));
        card.append(details);
        if(!c.analyzed_at)card.append(el('p','Noch nicht im Detail analysiert. Bitte oben „Buchungsdetails & Preise auslesen“ wählen.'));
        if(d.note)card.append(el('small',d.note));
        const exp=document.createElement('details');exp.style.marginTop='12px';
        const summary=el('summary','📩 Original-E-Mail ansehen');summary.style.cursor='pointer';exp.append(summary);
        const body=el('pre','Zum Anzeigen öffnen.');body.style.whiteSpace='pre-wrap';body.style.wordBreak='break-word';body.style.maxHeight='350px';body.style.overflow='auto';body.style.fontFamily='inherit';
        exp.append(body);
        exp.addEventListener('toggle',async()=>{if(!exp.open||exp.dataset.loaded)return;body.textContent='E-Mail wird aus Outlook geladen …';try{const m=await api('message_detail',{id:c.id});body.textContent=`Von: ${m.sender||c.sender||''}\nDatum: ${m.received_at||''}\nBetreff: ${m.subject||''}\n\n${m.body||'Kein Text verfügbar'}`;exp.dataset.loaded='1';}catch(e){body.textContent='E-Mail konnte nicht geladen werden: '+e.message;}});
        card.append(exp);
        if(isCancelled){card.append(el('p','Diese Stornierung wird nicht als aktive Buchung zugeordnet.'));}
        else {
          const match=tripMatch(d);
          if(match){card.append(el('p',`💡 Mögliche Reise: „${match.name}“ – bitte prüfen.`));card.append(btn('✓ Zu dieser Reise hinzufügen',()=>void assign(c,match).catch(e=>alert(e.message))));}
          if(state.trips.length)addSelect(c,card);
          card.append(btn('＋ Neue Reise erstellen',()=>void createTrip(c).catch(e=>alert(e.message))));
        }
        card.append(btn('Ausblenden',async()=>{try{await api('review',{id:c.id,status:'ignored'});await show();}catch(e){alert(e.message)}}));
        target.append(card);
      }
    }catch(e){target.replaceChildren(el('p','Fehler: '+e.message));}
  }
  async function analyze(){if(busy)return;busy=true;const b=document.querySelector('#outlookAnalyzeButton');if(b)b.disabled=true;
    try{let n=0;for(let i=0;i<10;i++){const r=await api('analyze');n+=r.analyzed||0;const s=document.querySelector('#outlookSmartProgress');if(s)s.textContent=`${n} E-Mails im Detail geprüft. ${r.remaining===0?'Fertig.':'Weitere folgen …'}`;if(!r.analyzed||r.remaining===0)break;}await show();}catch(e){alert('Analyse fehlgeschlagen: '+e.message);}finally{busy=false;if(b)b.disabled=false;}}
  function mount(){const panel=document.querySelector('#outlookImportPanel');if(!panel||panel===lastPanel)return;lastPanel=panel;
    const wrap=el('section');wrap.id='outlookSmart';wrap.style.marginTop='18px';
    wrap.append(el('h3','✨ Reisezuordnung mit einem Klick'));
    wrap.append(el('p','Zuerst E-Mails scannen, dann Buchungsdetails auswerten. Unklare Angaben bitte prüfen.'));
    const analyzeButton=btn('Buchungsdetails & Preise auslesen',()=>void analyze());analyzeButton.id='outlookAnalyzeButton';wrap.append(analyzeButton);
    wrap.append(btn('Vorschläge aktualisieren',()=>void show()));
    const p=el('p');p.id='outlookSmartProgress';wrap.append(p);
    const results=el('div');results.id='outlookSmartResults';wrap.append(results);
    wrap.append(btn('🧹 Alte Vorschläge bereinigen',async()=>{
      if(!confirm('Alle noch nicht übernommenen Outlook-Vorschläge entfernen und Scan zurücksetzen?'))return;
      try{await api('clean_and_rescan');await show();alert('Offene Vorschläge bereinigt. Jetzt oben den Posteingang erneut scannen.');}
      catch(e){alert('Bereinigung fehlgeschlagen: '+e.message);}
    }));
    panel.append(wrap);void show();
  }
  document.addEventListener("myjourney:outlook-updated", () => void show());
  new MutationObserver(mount).observe(document.body,{subtree:true,childList:true});mount();
})();
