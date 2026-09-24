import {CONFIG} from "./config.js";
const fleet=[
{id:"5T-01",type:"5 Ton",capacity:5,color:"#4ea1ff"},{id:"5T-02",type:"5 Ton",capacity:5,color:"#6d7cff"},{id:"5T-03",type:"5 Ton",capacity:5,color:"#9a72ff"},{id:"5T-04",type:"5 Ton",capacity:5,color:"#48c7d9"},
{id:"3T-01",type:"3 Ton",capacity:3,color:"#55c978"},{id:"3T-02",type:"3 Ton",capacity:3,color:"#9bd24f"},
{id:"VAN-01",type:"Sprinter",capacity:null,color:"#e7b84d"},{id:"VAN-02",type:"Sprinter",capacity:null,color:"#f08a4b"}];
let jobs=JSON.parse(localStorage.getItem("ms_jobs")||"[]");
document.querySelector("#app").innerHTML=`<header class="header"><div class="brandmark"><span class="brand-dot"></span><div><b>MOULDING</b><small>ROUTE CONTROL</small></div></div><div class="header-actions"><span class="live-dot"></span><span class="live-label">LIVE</span><button class="icon-btn" id="export" title="Export">↓</button><button class="btn primary" id="opt">Optimize</button></div></header><main class="main"><aside class="side"><div class="hero-row"><div><span class="eyebrow">TODAY</span><h1>Delivery board</h1></div><div class="date-pill">8 FLEET</div></div><div class="stats"><div class="stat"><span>Stops</span><b id="jc">0</b></div><div class="stat"><span>Assigned</span><b id="ac">0</b></div><div class="stat"><span>Vehicles</span><b>8</b></div></div><section class="card add-card"><div class="card-top"><div><span class="eyebrow">QUICK ADD</span><b>New delivery</b></div><span class="plus-mark">+</span></div><form id="form"><div class="grid"><div class="field"><label>Customer</label><input id="customer" required placeholder="Customer name"></div><div class="field"><label>Vehicle</label><select id="vehicle"><option>5 Ton</option><option>3 Ton</option><option>Sprinter</option></select></div><div class="field full"><label>Delivery address</label><input id="address" required placeholder="House number + street"></div><div class="field"><label>Service</label><input id="service" type="number" value="15" min="0"></div><div class="field"><label>Load / tons</label><input id="weight" type="number" step=".1" min="0" placeholder="Optional"></div></div><button class="btn primary save-btn">Add to board <span>→</span></button></form></section><section class="card routes-card"><div class="card-top"><div><span class="eyebrow">OPERATIONS</span><b>Routes</b></div><span class="muted">Drag to rebalance</span></div><div id="routes"></div></section></aside><section class="map"><div class="map-toolbar"><div class="map-title"><span class="eyebrow">LIVE MAP</span><b>Fleet overview</b></div><div class="map-chips"><div class="chip">5T <strong>4</strong></div><div class="chip">3T <strong>2</strong></div><div class="chip">VAN <strong>2</strong></div></div></div><div id="map"></div></section></main>`;
const map=L.map("map").setView([51.0447,-114.0719],10);L.tileLayer(CONFIG.mapTiles,{attribution:CONFIG.mapAttribution}).addTo(map);let layers=[];
const save=()=>localStorage.setItem("ms_jobs",JSON.stringify(jobs));
const esc=x=>String(x??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

// No-API-key Alberta address autocomplete using Photon/OpenStreetMap.
// Photon supports search-as-you-type, house/street layers, country filtering and bounding boxes.
const addressAC={timer:null,controller:null,results:[],active:-1};
function clearAddressSuggestions(){const box=document.querySelector("#address-suggestions");if(box){box.innerHTML="";box.classList.remove("show")}addressAC.results=[];addressAC.active=-1}
function addressText(item){const p=item.properties||{}; const house=p.housenumber||p.house_number||""; const street=p.street||p.road||p.name||""; const first=[house,street].filter(Boolean).join(" "); return first || p.label || item.display_name || "Address result"}
function addressDetail(item){const p=item.properties||{};return [p.city||p.locality||p.district,p.state||p.province||"Alberta",p.postcode].filter(Boolean).join(", ")}
function renderAddressSuggestions(results,message=""){
  const box=document.querySelector("#address-suggestions");if(!box)return;
  box.innerHTML="";addressAC.results=results;addressAC.active=-1;
  if(!results.length){
    if(message){const d=document.createElement("div");d.className="address-no-results";d.textContent=message;box.appendChild(d);box.classList.add("show");}
    else box.classList.remove("show");
    return;
  }
  results.forEach((item,i)=>{
    const b=document.createElement("button");b.type="button";b.className="address-suggestion";b.setAttribute("role","option");
    const title=addressText(item)||item.display_name||"Address";const detail=addressDetail(item);
    b.innerHTML=`<span class="address-suggestion-icon">⌖</span><span class="address-suggestion-copy"><strong>${esc(title)}</strong><small>${esc(detail || "Alberta, Canada")}${item.streetOnly?' • street match — verify house number':''}</small></span><span class="address-chevron">›</span>`;
    b.addEventListener("mousedown",e=>{e.preventDefault();selectAddressSuggestion(i)});box.appendChild(b)
  });box.classList.add("show")
}
function photonFeatureToAddress(f,streetOnly=false){
  const p=f.properties||{},coords=f.geometry&&f.geometry.coordinates||[];
  const label=[p.housenumber,p.street||p.name,p.city||p.locality,p.state,p.postcode,p.country].filter(Boolean).join(", ");
  return {display_name:label,lat:Number(coords[1]),lon:Number(coords[0]),place_id:p.osm_id||label,provider:"photon",photon:f,streetOnly}
}
function isAlberta(p){
  const country=String(p.countrycode||p.country_code||"").toLowerCase();
  const state=String(p.state||p.province||"").toLowerCase();
  return (country==="ca"||country==="canada")&&(state==="alberta"||state==="ab");
}
async function photonSearch(q,extra={},endpoint=CONFIG.geocoder){
  if(addressAC.controller) addressAC.controller.abort();
  addressAC.controller=new AbortController();
  const params=new URLSearchParams({q,limit:"12",lang:"en",countrycode:"CA",bbox:CONFIG.geocoderBbox,...extra});
  const r=await fetch(`${endpoint}?${params}`,{signal:addressAC.controller.signal,headers:{Accept:"application/json"}});
  if(!r.ok)throw new Error(`Address service returned ${r.status}`);
  const data=await r.json();return Array.isArray(data.features)?data.features:[];
}
function filterAlberta(features){return features.filter(f=>isAlberta(f.properties||{}));}
function hasHouse(p){return /^\\d/.test(String(p.housenumber||'').trim());}
async function fetchAlbertaAddresses(q){
  if(!/\\d/.test(q)) return {results:[],message:"Start with the house number, then the street name."};
  const normalized=q.replace(/\\s+/g," ").trim();
  let features=[];
  // Photon/OpenStreetMap house-level search. Alberta is enforced by bbox + client-side validation.
  try{features=await photonSearch(normalized,{layer:"house"});}catch(e){if(e.name==="AbortError")throw e;}
  let results=filterAlberta(features).filter(f=>hasHouse(f.properties||{})).map(f=>photonFeatureToAddress(f,false));
  // Structured Photon search is useful when the user types house number + street + city.
  if(!results.length){
    const m=normalized.match(/^([0-9]+[A-Za-z-]*)\\s+(.+)$/);
    if(m){
      try{
        const endpoint=CONFIG.geocoder.replace(/\\/api\\/?$/,'/structured');
        const params=new URLSearchParams({housenumber:m[1],street:m[2],state:"Alberta",countrycode:"CA",limit:"12",lang:"en",bbox:CONFIG.geocoderBbox});
        if(addressAC.controller) addressAC.controller.abort();
        addressAC.controller=new AbortController();
        const r=await fetch(`${endpoint}?${params}`,{signal:addressAC.controller.signal,headers:{Accept:"application/json"}});
        if(r.ok){const d=await r.json();features=Array.isArray(d.features)?d.features:[];results=filterAlberta(features).filter(f=>hasHouse(f.properties||{})).map(f=>photonFeatureToAddress(f,false));}
      }catch(e){if(e.name==="AbortError")throw e;}
    }
  }
  // Broader Photon search catches OSM buildings/places that aren't indexed as house layer.
  if(!results.length){
    try{features=await photonSearch(normalized);}catch(e){if(e.name==="AbortError")throw e;features=[];}
    results=filterAlberta(features).filter(f=>hasHouse(f.properties||{})).map(f=>photonFeatureToAddress(f,false));
  }
  if(results.length)return {results:results.slice(0,8)};
  return {results:[],message:"No exact Alberta address found. Try adding the city or postal code."};
}
async function resolveAddressCoordinates(text){
  try{
    const fs=await photonSearch(text,{layer:"house"});
    const f=filterAlberta(fs).find(x=>hasHouse(x.properties||{})&&Number.isFinite(Number(x.geometry?.coordinates?.[1]))&&Number.isFinite(Number(x.geometry?.coordinates?.[0])));
    if(f)return {lat:Number(f.geometry.coordinates[1]),lng:Number(f.geometry.coordinates[0])};
  }catch(e){}
  return null;
}
function selectAddressSuggestion(i){
  const item=addressAC.results[i];if(!item)return;const input=document.querySelector("#address");
  let value=item.display_name||"";
  if(item.streetOnly){const typed=input.value.trim();const m=typed.match(/^\s*([0-9]+[A-Za-z-]*)\s+/);if(m)value=`${m[1]} ${addressText(item)}`;}
  input.value=value;input.dataset.lat=Number.isFinite(item.lat)?item.lat:"";input.dataset.lng=Number.isFinite(item.lon)?item.lon:"";input.dataset.placeId=item.place_id||"";input.dataset.provider="photon";input.dataset.selected="1";input.dataset.streetOnly=item.streetOnly?"1":"";clearAddressSuggestions();
}
function setupAddressAutocomplete(){
  const input=document.querySelector("#address");if(!input)return;input.setAttribute("autocomplete","off");input.setAttribute("aria-autocomplete","list");input.setAttribute("placeholder","House number + street (e.g. 123 17 Ave SW)");input.parentElement.classList.add("address-autocomplete-wrap");
  const box=document.createElement("div");box.id="address-suggestions";box.className="address-suggestions";box.setAttribute("role","listbox");input.parentElement.appendChild(box);
  input.addEventListener("input",()=>{input.dataset.selected="";input.dataset.lat="";input.dataset.lng="";input.dataset.streetOnly="";const q=input.value.trim();clearTimeout(addressAC.timer);if(q.length<3){clearAddressSuggestions();return}addressAC.timer=setTimeout(async()=>{try{const out=await fetchAlbertaAddresses(q);renderAddressSuggestions(out.results,out.message)}catch(e){if(e.name!=="AbortError")renderAddressSuggestions([],"Address search is temporarily unavailable. You can try again in a moment.")}},300)});
  input.addEventListener("keydown",e=>{if(!box.classList.contains("show")||!addressAC.results.length)return;if(e.key==="ArrowDown"){e.preventDefault();addressAC.active=Math.min(addressAC.active+1,addressAC.results.length-1)}else if(e.key==="ArrowUp"){e.preventDefault();addressAC.active=Math.max(addressAC.active-1,0)}else if(e.key==="Enter"&&addressAC.active>=0){e.preventDefault();selectAddressSuggestion(addressAC.active);return}else if(e.key==="Escape"){clearAddressSuggestions();return}else return;box.querySelectorAll(".address-suggestion").forEach((x,i)=>x.classList.toggle("active",i===addressAC.active))});
  document.addEventListener("click",e=>{if(!input.parentElement.contains(e.target))clearAddressSuggestions()})
}
setupAddressAutocomplete();

document.querySelector("#form").onsubmit=async e=>{e.preventDefault();const a=document.querySelector("#address");if(!a.dataset.selected || !/\d+/.test(a.value)){alert("Please select an Alberta address suggestion first.");a.focus();return}const submit=e.submitter;submit.disabled=true;submit.textContent="Saving…";let lat=a.dataset.lat?+a.dataset.lat:null,lng=a.dataset.lng?+a.dataset.lng:null;if(!Number.isFinite(lat)||!Number.isFinite(lng)||a.dataset.streetOnly){const p=await resolveAddressCoordinates(a.value);if(p){lat=p.lat;lng=p.lng}else if(a.dataset.streetOnly){alert("That street was found, but an exact house location was not available. Please try adding the city or postal code.");submit.disabled=false;submit.textContent="+ Save Delivery";return}}jobs.push({id:crypto.randomUUID(),customer:document.querySelector("#customer").value.trim(),address:a.value.trim(),vehicle:document.querySelector("#vehicle").value,service:Number(document.querySelector("#service").value||15),weight:Number(document.querySelector("#weight").value||0),truckId:null,lat,lng});save();e.target.reset();document.querySelector("#service").value=15;submit.disabled=false;submit.textContent="+ Save Delivery";render()};
document.querySelector("#opt").onclick=()=>{for(const type of ["5 Ton","3 Ton","Sprinter"]){let ts=fleet.filter(t=>t.type===type),list=jobs.filter(j=>j.vehicle===type);list.forEach((j,i)=>j.truckId=ts[i%ts.length].id)}save();render()};
document.querySelector("#export").onclick=()=>{let rows=[["ID","Customer","Address","Vehicle","Truck","Stop","ServiceMinutes","WeightTons"]];fleet.forEach(t=>jobs.filter(j=>j.truckId===t.id).forEach((j,i)=>rows.push([j.id,j.customer,j.address,j.vehicle,t.id,i+1,j.service,j.weight])));jobs.filter(j=>!j.truckId).forEach(j=>rows.push([j.id,j.customer,j.address,j.vehicle,"","",j.service,j.weight]));let a=document.createElement("a");a.href=URL.createObjectURL(new Blob([rows.map(r=>r.map(v=>`"${String(v).replaceAll('"','""')}"`).join(",")).join("\\n")],{type:"text/csv"}));a.download="moulding-routes.csv";a.click()};
function render(){document.querySelector("#jc").textContent=jobs.length;document.querySelector("#ac").textContent=jobs.filter(j=>j.truckId).length;let root=document.querySelector("#routes");root.innerHTML="";fleet.forEach(t=>{let stops=jobs.filter(j=>j.truckId===t.id),el=document.createElement("div");el.className="truck";el.innerHTML=`<div class="truckhead"><div class="truck-name"><span class="truck-dot" style="background:${t.color}"></span><div><b>${t.id}</b><div class="muted">${t.type} · ${t.capacity?t.capacity+"T":"Van"} · ${stops.length} stops</div></div></div><button class="share-btn" data-share="${t.id}">Share ↗</button></div><div class="truckbody" data-truck="${t.id}"></div>`;let body=el.querySelector(".truckbody");body.ondragover=e=>{e.preventDefault();body.classList.add("over")};body.ondragleave=()=>body.classList.remove("over");body.ondrop=e=>{e.preventDefault();body.classList.remove("over");move(e.dataTransfer.getData("job"),t.id)};if(!stops.length)body.innerHTML='<div class="empty">Drop compatible job here</div>';stops.forEach((j,i)=>{let s=document.createElement("div");s.className="stop";s.draggable=true;s.innerHTML=`<div class="num">${i+1}</div><div class="info"><b>${esc(j.customer)}</b><span>${esc(j.address)}</span></div><span class="pill">${j.vehicle}</span>`;s.ondragstart=e=>e.dataTransfer.setData("job",j.id);body.appendChild(s)});root.appendChild(el)});root.querySelectorAll("[data-share]").forEach(b=>b.onclick=()=>share(b.dataset.share));drawMap()}
function move(id,tid){let j=jobs.find(x=>x.id===id),t=fleet.find(x=>x.id===tid);if(j&&t&&j.vehicle===t.type){j.truckId=tid;save();render()}}
async function geo(q){try{const params=new URLSearchParams({q:`${q}, Alberta, Canada`,limit:"5",lang:"en",bbox:CONFIG.geocoderBbox});const r=await fetch(`${CONFIG.geocoder}?${params}`,{headers:{Accept:"application/json"}});if(!r.ok)return null;const d=await r.json();const x=(d.features||[]).map(photonFeatureToAddress).find(v=>{const p=v.photon?.properties||{};const state=String(p.state||p.province||"").toLowerCase();const country=String(p.countrycode||p.country_code||p.country||"").toLowerCase();return (country==="ca"||country==="canada")&&(state==="alberta"||state==="ab")&&Number.isFinite(v.lat)&&Number.isFinite(v.lon)});return x?{lat:x.lat,lng:x.lon}:null}catch{return null}}
async function drawMap(){layers.forEach(x=>map.removeLayer(x));layers=[];let all=[];for(const t of fleet){let stops=jobs.filter(j=>j.truckId===t.id),pts=[];for(let i=0;i<stops.length;i++){let j=stops[i];if(j.lat==null){let p=await geo(j.address);if(p){Object.assign(j,p);save()}}if(j.lat==null)continue;pts.push([j.lat,j.lng]);all.push([j.lat,j.lng]);let icon=L.divIcon({className:"",html:`<div class="map-marker" style="background:${t.color}">${i+1}</div>`,iconSize:[29,29],iconAnchor:[14,14]});layers.push(L.marker([j.lat,j.lng],{icon}).bindPopup(`<b>${esc(j.customer)}</b><br>${esc(j.address)}<br>${t.id}`).addTo(map))}if(pts.length>1){let line=L.polyline(pts,{color:t.color,weight:5,opacity:.8});line.on("mouseover",()=>line.setStyle({weight:10}));line.on("mouseout",()=>line.setStyle({weight:5}));layers.push(line.addTo(map))}}if(all.length)map.fitBounds(all,{padding:[35,35],maxZoom:13})}
async function share(id){let t=fleet.find(x=>x.id===id),stops=jobs.filter(j=>j.truckId===id);for(let j of stops)if(j.lat==null){let p=await geo(j.address);if(p)Object.assign(j,p)}let c=stops.filter(j=>j.lat!=null).map(j=>`${j.lat},${j.lng}`);if(!c.length)return alert("Could not geocode the route.");let u=`https://www.google.com/maps/dir/?api=1&origin=${encodeURIComponent(c[0])}&destination=${encodeURIComponent(c.at(-1))}&travelmode=driving`;if(c.length>2)u+=`&waypoints=${encodeURIComponent(c.slice(1,-1).join("|"))}`;if(navigator.share)navigator.share({title:t.id+" Route",url:u}).catch(()=>{});else navigator.clipboard.writeText(u).then(()=>alert("Google Maps link copied."))}
render();if("serviceWorker"in navigator)navigator.serviceWorker.register("./sw.js").catch(()=>{});