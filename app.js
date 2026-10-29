
const KEY='myjourney.pwa.v1';

const demo={
  trips:[{
    id:'demo',
    name:'Malaysia & Thailand',
    start:'2027-01-21',
    end:'2027-02-09',
    destination:'Malaysia & Thailand'
  }],
  items:[
    {id:'a',tripId:'demo',type:'flight',title:'Anreise nach Kuala Lumpur',start:'2027-01-21T10:00',end:'',location:'Kuala Lumpur International Airport',address:'',lat:2.7456,lng:101.7072,confirmation:'',cost:0,notes:'Beispiel – bitte durch deine tatsächlichen Flugdaten ersetzen',source:'demo'},
    {id:'b',tripId:'demo',type:'hotel',title:'Hotel in Kuala Lumpur',start:'2027-01-22T15:00',end:'2027-01-25T11:00',location:'Kuala Lumpur',address:'',lat:3.139,lng:101.687,confirmation:'',cost:0,notes:'Beispielunterkunft, keine echte Buchung',source:'demo'}
  ]
};

let state;
try {
  state=JSON.parse(localStorage.getItem(KEY))||structuredClone(demo);
} catch {
  state=structuredClone(demo);
}
if(!Array.isArray(state.trips)||!Array.isArray(state.items)){
  state=structuredClone(demo);
}

let selected=state.trips[0]?.id||null;
let tab='overview',editId=null,map=null;

const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({
  '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
}[c]));
const uid=()=>globalThis.crypto?.randomUUID?.()||`${Date.now()}-${Math.random()}`;

function save(){
  try{
    localStorage.setItem(KEY,JSON.stringify(state));
    $('#storageStatus').textContent='Lokal gespeichert';
  }catch{
    $('#storageStatus').textContent='Speicher voll / gesperrt';
  }
}

function fmt(s){
  if(!s)return '—';
  const d=new Date(s);
  return isNaN(d)?esc(s):new Intl.DateTimeFormat('de-DE',{
    day:'2-digit',month:'short',year:'numeric',
    hour:s.includes('T')?'2-digit':undefined,
    minute:s.includes('T')?'2-digit':undefined
  }).format(d);
}

function money(n){
  return new Intl.NumberFormat('de-DE',{
    style:'currency',currency:'EUR'
  }).format(Number(n)||0);
}

function items(){
  return state.items
    .filter(i=>i.tripId===selected)
    .sort((a,b)=>(a.start||'').localeCompare(b.start||''));
}

const types={
  flight:'✈️ Flug',
  hotel:'🛏️ Unterkunft',
  transfer:'🚆 Transfer',
  activity:'🎟️ Aktivität',
  other:'📌 Sonstiges'
};

function render(){
  const ts=$('#tripSelect');
  ts.innerHTML=state.trips.map(t=>
    `<option value="${esc(t.id)}">${esc(t.name)}</option>`
  ).join('');

  if(!state.trips.some(t=>t.id===selected)){
    selected=state.trips[0]?.id||null;
  }
  ts.value=selected||'';

  document.querySelectorAll('[data-tab]').forEach(b=>
    b.classList.toggle('active',b.dataset.tab===tab)
  );

  if(map){map.remove();map=null;}

  ({overview,bookings,plan,map:mapview,settings}[tab]||overview)();
}

function section(title,button=''){
  return `<div class="sectionhead"><h2>${title}</h2>${button}</div>`;
}

function entry(i,buttons=true){
  return `<article class="card entry">
    <div>
      <span class="tag">${types[i.type]||types.other}</span>
      <h3 style="margin:10px 0 4px">${esc(i.title)}</h3>
      <div class="muted">${fmt(i.start)}${i.end?' – '+fmt(i.end):''}</div>
      <p>📍 ${esc(i.location||'Ort nicht angegeben')}</p>
      ${i.confirmation?`<p>Buchungsnummer: ${esc(i.confirmation)}</p>`:''}
      ${i.notes?`<p>${esc(i.notes)}</p>`:''}
      ${i.cost?`<strong>${money(i.cost)}</strong>`:''}
    </div>
    ${buttons?`<div class="rowactions">
      <button class="secondary mini" data-edit="${esc(i.id)}">Bearbeiten</button>
      <button class="danger mini" data-delete="${esc(i.id)}">✕</button>
    </div>`:''}
  </article>`;
}

function empty(label){
  return `<div class="card empty">
    <div class="emoji">🧳</div><p>${label}</p>
  </div>`;
}

function overview(){
  const t=state.trips.find(t=>t.id===selected),a=items();
  $('#view').innerHTML=
    section('Deine Reise',`<button class="primary" data-new="item">+ Buchung</button>`)+
    (t?`<div class="grid">
      <div class="card">
        <div class="muted">Reiseziel</div>
        <div class="big" style="font-size:21px">${esc(t.destination||t.name)}</div>
      </div>
      <div class="card">
        <div class="muted">Reisezeitraum</div>
        <div class="big" style="font-size:17px">${fmt(t.start)} – ${fmt(t.end)}</div>
      </div>
      <div class="card">
        <div class="muted">Erfasste Kosten</div>
        <div class="big">${money(a.reduce((s,i)=>s+(Number(i.cost)||0),0))}</div>
        <div class="muted">${a.length} Einträge</div>
      </div>
    </div>`:empty('Lege deine erste Reise an.'))+
    section('Nächste Buchungen')+
    (a.length?`<div class="stack">${a.slice(0,4).map(i=>entry(i,false)).join('')}</div>`:
      empty('Noch keine Buchungen vorhanden.'));
}

function bookings(){
  const a=items();
  $('#view').innerHTML=
    section('Alle Buchungen','<button class="primary" data-new="item">+ Eintrag</button>')+
    `<div class="card" style="margin-bottom:14px">
      <strong>Import aus JSON-Datei</strong>
      <p class="muted">Importiere Buchungen aus einer strukturierten Datei.
      PDFs und E-Mails werden noch nicht automatisch erkannt.</p>
      <input type="file" id="jsonImport" accept=".json,application/json">
    </div>`+
    (a.length?'<div class="stack">'+a.map(i=>entry(i)).join('')+'</div>':
      empty('Hier erscheinen deine Flüge, Hotels und Aktivitäten.'));
}

function plan(){
  const groups={};
  items().forEach(i=>{
    const d=(i.start||'Ohne Datum').slice(0,10);
    (groups[d]??=[]).push(i);
  });
  const content=Object.entries(groups).map(([d,a])=>
    '<h3 class="day">'+fmt(d)+'</h3><div class="timeline">'+
    a.map(i=>entry(i)).join('')+'</div>'
  ).join('');

  $('#view').innerHTML=
    section('Dein Tagesplan','<button class="primary" data-new="item">+ Aktivität</button>')+
    (content||empty('Sobald du Einträge hinzufügst, entsteht automatisch dein Tagesplan.'));
}

function mapview(){
  const a=items().filter(i=>
    Number.isFinite(Number(i.lat))&&
    Number.isFinite(Number(i.lng))&&
    i.lat!==''&&i.lng!==''&&i.lat!=null&&i.lng!=null
  );

  $('#view').innerHTML=section('Kartenansicht')+
    `<div class="notice">Die Karte benötigt Internet. Standorte müssen als
    Koordinaten hinterlegt sein. Du kannst sie beim Bearbeiten einer Buchung
    ergänzen.</div><div id="map" class="mapbox"></div>`+
    (a.length?`<div class="stack" style="margin-top:16px">${
      a.map(i=>`<div class="card">${esc(i.title)} · ${esc(i.location||'')} ·
      <a target="_blank" rel="noopener"
      href="https://www.openstreetmap.org/?mlat=${encodeURIComponent(i.lat)}&mlon=${encodeURIComponent(i.lng)}#map=13/${encodeURIComponent(i.lat)}/${encodeURIComponent(i.lng)}">
      In OpenStreetMap öffnen</a></div>`).join('')
    }</div>`:'');

  if(typeof L==='undefined'){
    $('#map').innerHTML='<div class="empty">Kartenbibliothek offline nicht verfügbar. Verbinde dich mit dem Internet.</div>';
    return;
  }

  map=L.map('map').setView([a[0]?.lat||20,a[0]?.lng||10],a.length?8:2);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{
    attribution:'© OpenStreetMap-Mitwirkende',maxZoom:18
  }).addTo(map);

  if(a.length){
    const markers=a.map(i=>
      L.marker([Number(i.lat),Number(i.lng)]).addTo(map).bindPopup(esc(i.title))
    );
    if(markers.length>1){
      map.fitBounds(L.featureGroup(markers).getBounds().pad(.2));
    }else{
      map.setView([Number(a[0].lat),Number(a[0].lng)],12);
    }
  }
}

function settings(){
  $('#view').innerHTML=section('Daten & Einstellungen')+
    `<div class="stack">
      <div class="card">
        <h3>Deine Reisen</h3>
        <p class="muted">Reisen verwalten und sichern</p>
        <div class="rowactions">
          <button class="primary" data-new="trip">+ Neue Reise</button>
          <button class="secondary" id="editTrip">Reise bearbeiten</button>
          <button class="danger" id="deleteTrip">Reise löschen</button>
        </div>
      </div>
      <div class="card">
        <h3>Datensicherung</h3>
        <p class="muted">Exportiere deine Reisedaten als JSON-Datei.</p>
        <button class="secondary" id="export">Alle Daten exportieren</button>
        <label class="secondary" style="display:inline-block;cursor:pointer">
          Backup wiederherstellen
          <input id="restore" type="file" accept=".json" hidden>
        </label>
      </div>
      <div class="card">
        <h3>🔗 Meine Verbindungen</h3>
        <p class="muted">Verbinde zukünftig dein Outlook-Konto,
        damit MyJourney Buchungsbestätigungen automatisch erkennen kann.</p>
        <button class="primary" id="connectOutlook" type="button">
          ✉️ Outlook verbinden
        </button>
        <p class="muted" style="margin-top:10px">
          Die Verbindung ist noch nicht eingerichtet.
          Der automatische E-Mail-Abruf folgt in einem späteren Schritt.
        </p>
      </div>
    </div>`;
}

const fields={
  trip:[
    ['name','Reisename','text',true],
    ['destination','Reiseziel','text'],
    ['start','Startdatum','date',true],
    ['end','Enddatum','date',true]
  ],
  item:[
    ['type','Kategorie','select',true],
    ['title','Bezeichnung','text',true],
    ['start','Beginn','datetime-local',true],
    ['end','Ende','datetime-local'],
    ['location','Ort','text'],
    ['address','Adresse','text'],
    ['lat','Breitengrad (optional)','number'],
    ['lng','Längengrad (optional)','number'],
    ['confirmation','Buchungsnummer','text'],
    ['cost','Kosten in EUR','number'],
    ['notes','Notizen','textarea']
  ]
};

function openEditor(kind,id=null){
  if(kind==='item'&&!selected){
    alert('Bitte zuerst eine Reise anlegen.');
    return;
  }

  editId=id;
  const data=id?(kind==='trip'?state.trips:state.items).find(x=>x.id===id):{};
  $('#dialogTitle').textContent=(id?'Bearbeiten: ':'Neu: ')+
    (kind==='trip'?'Reise':'Buchung / Aktivität');
  $('#editorForm').dataset.kind=kind;

  $('#fields').innerHTML=`<div class="formgrid">${
    fields[kind].map(([name,label,type,required])=>
      `<div class="field ${name==='notes'?'full':''}">
        <label for="f_${name}">${label}</label>
        ${type==='select'?
          `<select id="f_${name}" name="${name}">${
            Object.entries(types).map(([v,t])=>
              `<option value="${v}" ${data?.[name]===v?'selected':''}>${t}</option>`
            ).join('')
          }</select>`:
          type==='textarea'?
          `<textarea id="f_${name}" name="${name}" rows="3">${esc(data?.[name])}</textarea>`:
          `<input id="f_${name}" name="${name}" type="${type}"
          ${type==='number'?'step="any"':''}
          value="${esc(data?.[name])}" ${required?'required':''}>`
        }
      </div>`
    ).join('')
  }</div>`;

  $('#editor').showModal();
}

$('#editorForm').addEventListener('submit',e=>{
  e.preventDefault();
  const kind=e.currentTarget.dataset.kind;
  const values=Object.fromEntries(new FormData(e.currentTarget).entries());

  if(values.end&&values.start&&values.end<values.start){
    alert('Das Enddatum muss nach dem Beginn liegen.');
    return;
  }

  if(kind==='item'){
    values.cost=Number(values.cost)||0;
    for(const n of ['lat','lng']){
      values[n]=values[n]===''?null:Number(values[n]);
    }
  }

  const arr=kind==='trip'?state.trips:state.items;
  if(editId){
    Object.assign(arr.find(x=>x.id===editId),values);
  }else{
    const obj={
      ...values,
      id:uid(),
      ...(kind==='item'?{tripId:selected,source:'manual'}:{})
    };
    arr.push(obj);
    if(kind==='trip')selected=obj.id;
  }

  save();
  $('#editor').close();
  render();
});

$('#closeDialog').onclick=$('#cancelDialog').onclick=()=>$('#editor').close();
$('#addTrip').onclick=()=>openEditor('trip');

$('#tripSelect').onchange=e=>{
  selected=e.target.value;
  render();
};

document.querySelectorAll('[data-tab]').forEach(b=>
  b.onclick=()=>{
    tab=b.dataset.tab;
    render();
  }
);

$('#view').addEventListener('click',e=>{
  const b=e.target.closest('button');
  if(!b)return;

  if(b.dataset.new)openEditor(b.dataset.new);
  if(b.dataset.edit)openEditor('item',b.dataset.edit);

  if(b.dataset.delete&&confirm('Eintrag wirklich löschen?')){
    state.items=state.items.filter(i=>i.id!==b.dataset.delete);
    save();
    render();
  }

  if(b.id==='editTrip')openEditor('trip',selected);

  if(b.id==='deleteTrip'&&confirm('Diese Reise mit allen Buchungen löschen?')){
    state.trips=state.trips.filter(t=>t.id!==selected);
    state.items=state.items.filter(i=>i.tripId!==selected);
    selected=state.trips[0]?.id||null;
    save();
    render();
  }

  if(b.id==='connectOutlook'){
    alert('Die Microsoft-Anmeldung wird als Nächstes eingerichtet. Dein Outlook-Postfach ist noch nicht verbunden.');
  }

  if(b.id==='export'){
    const blob=new Blob([JSON.stringify(state,null,2)],{
      type:'application/json'
    });
    const url=URL.createObjectURL(blob);
    const a=document.createElement('a');
    a.href=url;
    a.download='myjourney-backup.json';
    a.click();
    setTimeout(()=>URL.revokeObjectURL(url),1000);
  }
});

$('#view').addEventListener('change',async e=>{
  if(!['jsonImport','restore'].includes(e.target.id))return;

  const file=e.target.files?.[0];
  if(!file)return;

  try{
    const data=JSON.parse(await file.text());

    if(e.target.id==='restore'){
      if(!Array.isArray(data.trips)||!Array.isArray(data.items)){
        throw Error('Ungültiges Backup');
      }
      if(!confirm('Alle bisherigen Daten durch das Backup ersetzen?'))return;
      state=data;
      selected=state.trips[0]?.id||null;
    }else{
      const arr=Array.isArray(data)?data:
        Array.isArray(data.bookings)?data.bookings:
        Array.isArray(data.items)?data.items:null;

      if(!arr)throw Error('Erwartet wird ein Array oder ein Objekt mit "bookings".');

      let count=0;
      for(const i of arr){
        if(!i||typeof i!=='object'||!i.title||!i.start)continue;
        const type=types[i.type]?i.type:'other';

        state.items.push({
          id:uid(),
          tripId:selected,
          type,
          title:String(i.title),
          start:String(i.start),
          end:String(i.end||''),
          location:String(i.location||''),
          address:String(i.address||''),
          lat:i.lat==null?null:Number(i.lat),
          lng:i.lng==null?null:Number(i.lng),
          confirmation:String(i.confirmation||''),
          cost:Number(i.cost)||0,
          notes:String(i.notes||''),
          source:'json'
        });
        count++;
      }
      alert(`${count} Einträge importiert.`);
    }

    save();
    render();
  }catch(err){
    alert('Import fehlgeschlagen: '+err.message);
  }
});

save();
render();

if('serviceWorker'in navigator&&location.protocol.startsWith('http')){
  navigator.serviceWorker.register('./sw.js').catch(()=>{});
}
