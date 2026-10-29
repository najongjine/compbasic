import fs from 'node:fs';
import vm from 'node:vm';
import assert from 'node:assert/strict';
const html=fs.readFileSync(new URL('index.html',import.meta.url),'utf8');
const raw=fs.readFileSync(new URL('suno_simple.txt',import.meta.url),'utf8');
const scripts=[...html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/g)];
const data=JSON.parse(scripts.find(s=>s[1].includes('application/json'))[2]);
const code=scripts.find(s=>!s[1].includes('application/json'))[2];
new vm.Script(code);
const original=raw.split(/\r?\n/).filter(l=>l.includes('#')).map(l=>l.trim().split(/\s*#\s*/).map(s=>s.trim()));
assert.equal(data.rows.length,329);
assert.deepEqual(data.rows.map(r=>[r.tag,r.ko]),original);
for(const row of data.rows){assert.ok(row.note);const s=data.songs[row.song];assert.ok(s?.artist&&s.title);if(s.video)assert.match(s.video,/^[\w-]{11}$/)}
class Element{
  constructor(){this.innerHTML='';this.textContent='';this.value='';this.hidden=false;this.listeners={};this.attrs={};this.dataset={};this.classList={toggle(){},remove(){}};this.style={};}
  addEventListener(k,fn){this.listeners[k]=fn}
  setAttribute(k,v){this.attrs[k]=v}
  scrollIntoView(){}focus(){}select(){}append(){}remove(){}
}
const els=new Map();
const get=id=>{if(!els.has(id))els.set(id,new Element());return els.get(id)};
get('dictionary-data').textContent=JSON.stringify(data);
let timerId=0;const timers=new Map(),intervals=new Map(),starts=[];
const document={getElementById:get,querySelectorAll:()=>[],body:new Element(),hidden:false,listeners:{},addEventListener(k,fn){this.listeners[k]=fn},createElement:()=>new Element(),execCommand:()=>true};
const context=vm.createContext({document,window:{addEventListener(){}},navigator:{clipboard:{writeText:async t=>{context.copied=t}}},setTimeout:(fn,ms)=>{const id=++timerId;timers.set(id,{fn,ms});return id},clearTimeout:id=>timers.delete(id),setInterval:(fn,ms)=>{const id=++timerId;intervals.set(id,{fn,ms});return id},clearInterval:id=>intervals.delete(id),console});
const run=s=>vm.runInContext(s,context);
run(code);
assert.equal((get('results').innerHTML.match(/class="tag-card/g)||[]).length,58);
run("activeCategory='전체';render()");
assert.equal((get('results').innerHTML.match(/class="tag-card/g)||[]).length,329);
run("query='가야금';render()");
assert.ok(get('results').innerHTML.includes('Gayageum'));
assert.ok(!get('results').innerHTML.includes('class="tag-card" data-id="0"'));
run("query='Nessun dorma';render()");assert.ok(get('results').innerHTML.includes('Luciano Pavarotti'));
run("query='<script>invalid';render()");assert.ok(get('results').innerHTML.includes('검색 초기화'));
run("query='';activeCategory='장르';activeGroup='한국 음악 장르';render()");
assert.equal((get('results').innerHTML.match(/class="tag-card/g)||[]).length,12);
const style=data.rows.find(r=>r.tag==='Piano');
const structure=data.rows.find(r=>r.tag==='[Intro]');
run(`toggleTag(${style.id});toggleTag(${structure.id})`);
assert.equal(get('style-tags').value,'Piano');assert.equal(get('structure-tags').value,'[Intro]');assert.equal(get('composer').hidden,false);
await get('copy-selection').listeners.click();assert.equal(context.copied,'스타일\nPiano\n\n노래 구조\n[Intro]');
run(`toggleTag(${style.id});toggleTag(${structure.id})`);assert.equal(get('composer').hidden,true);
class FakeAudio{
  constructor(){this.currentTime=0;this.destination={}}
  async resume(){}
  createGain(){return {gain:{value:0,setValueAtTime(){},exponentialRampToValueAtTime(){}},connect(){},disconnect(){}}}
  createOscillator(){return {frequency:{value:0},connect(){},disconnect(){},start:t=>starts.push(t),stop(){}}}
}
context.window.AudioContext=FakeAudio;
await run('startAudio(120)');assert.equal(run('isPlaying'),true);assert.equal(starts[0],.06);
run('audioContext.currentTime=.5');[...intervals.values()][0].fn();assert.ok(Math.abs(starts[1]-starts[0]-.5)<1e-9);
run('stopAudio()');assert.equal(intervals.size,0);assert.equal(run('soundNodes.size'),0);assert.equal(run('pulseTimers.size'),0);
await run('startAudio(180)');const first=starts.at(-1);run('audioContext.currentTime=.85');[...intervals.values()][0].fn();assert.ok(Math.abs(starts.at(-1)-first-1/3)<1e-9);
document.listeners.keydown({key:'Escape'});assert.equal(run('isPlaying'),false);assert.equal(intervals.size,0);
assert.ok(!html.includes('/*__DATA__*/'));
console.log('PASS: 329 original entries preserved; every entry has a song and listening note; scripts parse; category/subgroup filters; global Korean/song search; empty results; separate style/structure selection; copy; 120/180 BPM scheduler; Escape cancellation.');
console.log('Browser layout and actual audio output were not verified.');
