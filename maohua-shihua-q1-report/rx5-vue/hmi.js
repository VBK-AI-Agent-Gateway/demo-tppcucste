/* RX5 洛神座舱 HMI · Vue3 单文件逻辑层
   对接预留: sendCmd() → vehicleBridge({service,op,value,origin}) → WS网关 → CAN/UDS */
const {createApp,reactive,computed}=Vue;
const car=reactive({
  page:'home', gear:'P', speed:0, range:486, fuel:62, fuel100:7.2, waterTemp:92,
  tire:[2.5,2.5,2.5,2.5], tireWarn:false,
  temp:24,aOn:false,wind:1,circ:'内',ac:false,fogF:false,fogR:false,heatD:0,heatP:0,
  locked:true,doorL:false,doorR:false,sunroof:false,light:'近光',mirror:'展开',mode:'标准',
  vol:60,playing:false,songI:0,
  songs:['夜空中最亮的星','平凡之路','晴天','海阔天空'],
  nav:false,dest:'',route:'推荐',voice:false,voiceTxt:'',
  dark:true,bright:60,wifi:false,bt:'ROEWE-RX5-3F2A',ota:'v3.0.1',
  clock:'--:--',weather:'☀ 26°/18°',net:'📶 4G',radarDist:null
});
function sendCmd(svc,op,val){console.log('[TX]',svc,op,val)/* TODO: 对接 vehicleBridge */}
let _t;
function toast(m){const t=document.getElementById('toast');
  if(!t)return;t.textContent=m;t.style.opacity=1;clearTimeout(_t);_t=setTimeout(()=>t.style.opacity=0,1800)}
let radarT=null;
function setGear(g){car.gear=g;sendCmd('can','GEAR',g);
  if(g==='D'){car.playing=false;toast('行驶中 · 视频/复杂娱乐已暂停')}
  if(g==='R'){let d=1.5;clearInterval(radarT);
    radarT=setInterval(()=>{d-=.18;car.radarDist=Math.max(.2,d).toFixed(1);if(d<=.2)clearInterval(radarT)},500)}
  else{clearInterval(radarT);car.radarDist=null}
  if(g==='P')toast('P挡 · 娱乐功能已解锁')}
function startVoice(){if(car.voice){car.voice=false;return}
  car.voice=true;car.voiceTxt='“你好斑马” · 聆听中 · 可见即可说';
  sendCmd('voice','SESSION_START','90s');
  setTimeout(()=>{if(!car.voice)return;
    car.voiceTxt='→ 已将空调调至 24°C · 风量 2';car.temp=24;car.wind=2;
    sendCmd('voice','RESULT','hvac.temp=24');
    setTimeout(()=>{car.voice=false;toast('语音会话结束')},2500)},1800)}
setInterval(()=>{const d=new Date();car.clock=d.toTimeString().slice(0,5)},1000);
setInterval(()=>{ // 模拟数据源(生产: CAN订阅)
  car.speed=Math.min(120,Math.max(0,Math.round(car.speed+(Math.random()-.5)*8)));
  car.tireWarn=car.tire.some(t=>t<2.0);
  car.net=navigator.onLine?'📶 4G':'📴 离线'},2000);

const Dock={props:['p'],emits:['go'],template:`
<div class="h-20 bg-card border-t border-line flex justify-around items-center pb-1 shrink-0 z-30">
 <button v-for="d in items" :key="d.k" @click="$emit('go',d.k)"
  class="tbtn flex-col rounded-xl px-1 flex items-center justify-center gap-0.5"
  :class="p===d.k?'text-accent':''">
  <span class="text-xl">{{d.i}}</span><span class="text-[10px]">{{d.t}}</span></button></div>`,
 setup(){return{items:[
  {k:'home',i:'🏠',t:'主页'},{k:'vehicle',i:'🚗',t:'车控'},{k:'air',i:'❄️',t:'空调'},
  {k:'media',i:'🎵',t:'影音'},{k:'cam',i:'📷',t:'360'},{k:'park',i:'🅿️',t:'泊车'},
  {k:'voice',i:'🎙',t:'语音'},{k:'settings',i:'⚙️',t:'设置'}]}}};

const Home={template:`
<div class="p-4 space-y-3 pop">
 <div v-if="car.gear==='D'" class="bg-card border border-accent/40 rounded-2xl p-4">
  <div class="text-xs text-muted mb-1">🧭 D挡 · 导航优先</div>
  <div v-if="car.nav" class="text-xl font-bold text-accent">→ {{car.dest}} · 12分钟 · 8.6km</div>
  <div v-else class="text-sm text-muted">未规划路线 <button class="tbtn text-accent" @click="car.page='nav'">去导航›</button></div>
 </div>
 <div class="bg-card border border-line rounded-2xl p-4">
  <div class="text-xs text-muted mb-2">荣威RX5 · 点击部位控制</div>
  <div class="relative h-40">
   <svg viewBox="0 0 320 150" class="w-full h-full">
    <path d="M45 100 L68 56 Q88 40 130 40 L205 40 Q243 44 262 62 L276 100 Z" fill="rgba(79,195,247,.06)" stroke="#4FC3F7" stroke-width="3"/>
    <path d="M95 58 L110 44 L198 44 L216 58 Z" fill="rgba(79,195,247,.10)" stroke="#4FC3F7" stroke-width="2"/>
    <rect x="150" y="35" width="58" height="6" rx="3" :fill="car.sunroof?'#4FC3F7':'rgba(79,195,247,.3)'"/>
    <circle cx="95" cy="102" r="15" fill="none" :stroke="car.locked?'#8A9199':'#4FC3F7'" stroke-width="3.5"/>
    <circle cx="240" cy="102" r="15" fill="none" :stroke="car.locked?'#8A9199':'#4FC3F7'" stroke-width="3.5"/>
   </svg>
   <button class="tbtn absolute left-[8%] top-[24%] rounded-full border border-dashed border-accent text-accent text-[11px]" @click="car.locked=!car.locked;sendCmd('body','LOCK',car.locked);toast(car.locked?'已上锁🔒':'已解锁🔓')">🚪🔒</button>
   <button class="tbtn absolute left-[46%] top-[2%] rounded-full border border-dashed border-accent text-accent text-[11px]" @click="car.sunroof=!car.sunroof;toast('天窗'+(car.sunroof?'开1/2':'关'))">🌤天窗</button>
   <button class="tbtn absolute right-[6%] top-[24%] rounded-full border border-dashed border-accent text-accent text-[11px]" @click="car.doorR=!car.doorR;toast('右门'+(car.doorR?'开':'关'))">🚪</button>
   <button class="tbtn absolute left-[18%] bottom-0 rounded-full border border-dashed border-warn text-warn text-[11px]" @click="toast('胎压 '+car.tire.join('/'))">🛞</button>
   <button class="tbtn absolute right-[18%] bottom-0 rounded-full border border-dashed border-accent text-accent text-[11px]" @click="toast('尾门开启')">尾门</button>
  </div>
 </div>
 <div class="grid grid-cols-4 gap-2 text-center">
  <div class="bg-card rounded-xl p-2 border border-line"><div class="text-[10px] text-muted">车速</div><b class="font-din text-lg">{{car.speed}}</b></div>
  <div class="bg-card rounded-xl p-2 border border-line"><div class="text-[10px] text-muted">油量</div><b class="text-lg">{{car.fuel}}%</b></div>
  <div class="bg-card rounded-xl p-2 border border-line"><div class="text-[10px] text-muted">续航</div><b class="text-lg">{{car.range}}</b></div>
  <div class="bg-card rounded-xl p-2 border border-line"><div class="text-[10px] text-muted">油耗</div><b class="text-lg">{{car.fuel100}}</b></div>
 </div>
 <div class="grid grid-cols-3 gap-2 text-center">
  <div class="bg-card rounded-xl p-3 border border-line"><div class="font-bold text-xl">{{car.clock}}</div><div class="text-[10px] text-muted">{{car.weather}}</div></div>
  <div class="bg-card rounded-xl p-3 border border-line"><div class="text-[10px] text-muted">胎压</div><b :class="car.tireWarn?'text-danger':'text-ok'">{{car.tire[0]}}</b></div>
  <div class="bg-card rounded-xl p-3 border border-line"><div class="text-[10px] text-muted">水温</div><b>{{car.waterTemp}}°</b></div>
 </div>
</div>`};

const NavPage={template:`
<div class="p-4 space-y-3 pop h-full flex flex-col">
 <div class="relative flex-1 rounded-2xl border border-line overflow-hidden min-h-[300px]"
  style="background:linear-gradient(135deg,#0E2338,#14304E 60%,#0E2338)">
  <div class="absolute inset-x-0 top-[28%] h-6 bg-white/5 rotate-2"></div>
  <div class="absolute inset-x-0 top-[60%] h-6 bg-white/5 -rotate-1 w-3/4"></div>
  <div class="absolute inset-0 flex flex-col items-center justify-center gap-2">
   <div class="text-lg">🗺️ 高德地图 H5 · JSAPI 插槽</div>
   <div v-if="car.nav" class="text-sm text-accent">→ {{car.dest}} · 车道级引导 · 路况畅通</div>
   <div v-else class="text-xs text-muted">联网加载 · 离线降级为路书模式</div>
  </div>
 </div>
 <div class="grid grid-cols-4 gap-2">
  <button class="tbtn bg-accent rounded-xl font-bold text-slate-900" @click="go()">🔍 搜索</button>
  <button class="tbtn bg-card2 rounded-xl" @click="poi('充电桩')">⚡ 充桩</button>
  <button class="tbtn bg-card2 rounded-xl" @click="poi('加油站')">⛽ 加油</button>
  <button class="tbtn bg-card2 rounded-xl" @click="poi('停车场')">🅿️ 停车</button>
 </div>
 <div class="bg-card rounded-xl border border-line p-1 grid grid-cols-3 text-sm w-fit">
  <button v-for="r in ['推荐','高速优先','省油']" :key="r" class="tbtn px-4 rounded-lg"
   :class="car.route===r?'bg-accent text-slate-900 font-bold':''" @click="car.route=r;toast('路线偏好 '+r)">{{r}}</button>
 </div>
</div>`,
 setup(){const go=()=>{car.nav=true;car.dest='公司';sendCmd('nav','START','公司');toast('已开始导航')};
  const poi=t=>{go();car.dest=t+'·第1个';toast('已搜索附近'+t)};return{car,go,poi}}};

const AirPage={template:`
<div class="p-4 space-y-3 pop">
 <div class="bg-card rounded-2xl border border-line p-6 text-center">
  <div class="text-6xl font-bold text-accent">{{car.temp}}°C</div>
  <div class="mt-4 grid grid-cols-3 gap-3">
   <button class="tbtn bg-card2 rounded-xl text-2xl" @click="car.temp>16&&car.temp--;sendCmd('hvac','SET_TEMP',car.temp)">−</button>
   <button class="tbtn bg-card2 rounded-xl" :class="car.aOn?'text-accent font-bold':''" @click="car.aOn=!car.aOn;sendCmd('hvac','POWER',car.aOn)">{{car.aOn?'开':'关'}}</button>
   <button class="tbtn bg-card2 rounded-xl text-2xl" @click="car.temp<30&&car.temp++;sendCmd('hvac','SET_TEMP',car.temp)">＋</button>
  </div>
 </div>
 <div class="bg-card rounded-2xl border border-line p-4 grid grid-cols-3 gap-2 text-sm">
  <button class="tbtn bg-card2 rounded-xl" @click="car.wind=car.wind%4+1">风量 {{car.wind}}</button>
  <button class="tbtn bg-card2 rounded-xl" :class="car.ac?'text-accent':''" @click="car.ac=!car.ac">AC</button>
  <button class="tbtn bg-card2 rounded-xl" @click="car.circ=car.circ==='内'?'外':'内'">{{car.circ}}循环</button>
  <button class="tbtn bg-card2 rounded-xl" :class="car.fogF?'text-warn':''" @click="car.fogF=!car.fogF">前除雾</button>
  <button class="tbtn bg-card2 rounded-xl" :class="car.fogR?'text-warn':''" @click="car.fogR=!car.fogR">后除雾</button>
  <button class="tbtn bg-card2 rounded-xl" @click="car.heatD=car.heatD==3?0:car.heatD+1">主驾加热{{car.heatD}}</button>
  <button class="tbtn bg-card2 rounded-xl" @click="car.heatP=car.heatP==3?0:car.heatP+1">副驾加热{{car.heatP}}</button>
 </div>
</div>`};

const Media={template:`
<div class="p-4 space-y-3 pop">
 <div v-if="car.gear==='D'" class="bg-card rounded-xl border border-warn/40 p-3 text-warn text-sm text-center">⚠ D挡 · 视频已暂停，仅音频可播</div>
 <div class="bg-card rounded-2xl border border-line p-5 text-center">
  <div class="text-lg font-bold">{{car.songs[car.songI]}}</div>
  <div class="mt-3 flex justify-center gap-4">
   <button class="tbtn bg-card2 rounded-full" @click="car.songI=(car.songI+3)%4">⏮</button>
   <button class="tbtn bg-accent rounded-full w-20 font-bold text-slate-900" @click="toggle">▶/⏸</button>
   <button class="tbtn bg-card2 rounded-full" @click="car.songI=(car.songI+1)%4">⏭</button>
  </div>
  <div class="mt-4 flex items-center gap-3 text-sm"><span class="text-muted">音量</span>
   <input type="range" min="0" max="100" v-model.number="car.vol" class="flex-1"><b class="w-8">{{car.vol}}</b></div>
 </div>
 <div class="bg-card rounded-2xl border border-line p-4">
  <div class="text-xs text-muted mb-2">播放列表 · 本地音频</div>
  <button v-for="(s,i) in car.songs" :key="i" class="tbtn w-full text-left px-4 rounded-xl flex justify-between"
   :class="i===car.songI?'text-accent':''" @click="car.songI=i;car.playing=true">{{s}}
   <span v-if="i===car.songI&&car.playing" class="text-[10px]">▸播放中</span></button>
 </div>
 <div class="bg-card rounded-2xl border border-line p-4">
  <div class="text-xs text-muted mb-2">蓝牙电话 · {{car.bt}}</div>
  <div class="grid grid-cols-3 gap-2">
   <button class="tbtn bg-accent rounded-xl font-bold text-slate-900" @click="toast('拨打 138****6688')">📞 拨号</button>
   <button class="tbtn bg-card2 rounded-xl" @click="toast('通讯录加载中')">👥 通讯录</button>
   <button class="tbtn bg-card2 rounded-xl" @click="toast('通话记录已打开')">🕘 通话记录</button>
  </div>
 </div>
</div>`,
 setup(){const toggle=()=>{if(car.gear==='D'){toast('D挡 · 仅播音频，视频暂停');return}
  car.playing=!car.playing};return{car,toggle}}};

const Vehicle={template:`
<div class="p-4 space-y-3 pop">
 <div class="bg-card rounded-2xl border border-line p-4 space-y-3 text-sm">
  <div class="flex justify-between items-center"><span>整车门锁 ({{car.locked?'已锁':'开'}})</span>
   <button class="tbtn bg-accent rounded-xl px-4 font-bold text-slate-900" @click="car.locked=!car.locked;sendCmd('body','LOCK',car.locked)">{{car.locked?'解锁🔓':'上锁🔒'}}</button></div>
  <div class="flex justify-between items-center"><span>天窗</span>
   <button class="tbtn bg-card2 rounded-xl px-4" @click="car.sunroof=!car.sunroof">{{car.sunroof?'开启中':'关闭'}}</button></div>
  <div class="flex justify-between items-center"><span>左前门</span>
   <button class="tbtn bg-card2 rounded-xl px-4" @click="car.doorL=!car.doorL">{{car.doorL?'开':'关'}}</button></div>
  <div class="flex justify-between items-center"><span>右前门</span>
   <button class="tbtn bg-card2 rounded-xl px-4" @click="car.doorR=!car.doorR">{{car.doorR?'开':'关'}}</button></div>
  <div class="flex justify-between items-center"><span>驾驶模式</span>
   <div class="bg-card2 rounded-xl p-1 grid grid-cols-3 text-xs">
    <button v-for="m in ['经济','标准','运动']" :key="m" class="tbtn px-4 rounded-lg"
     :class="car.mode===m?'bg-accent text-slate-900 font-bold':''" @click="car.mode=m;sendCmd('powertrain','MODE',m)">{{m}}</button></div></div>
  <div class="flex justify-between items-center"><span>车灯</span>
   <div class="bg-card2 rounded-xl p-1 grid grid-cols-3 text-xs">
    <button v-for="l in ['近光','远光','示宽']" :key="l" class="tbtn px-4 rounded-lg"
     :class="car.light===l?'bg-accent text-slate-900':''" @click="car.light=l;toast(l+'已开启')">{{l}}</button></div></div>
  <div class="flex justify-between items-center"><span>后视镜</span>
   <div class="bg-card2 rounded-xl p-1 grid grid-cols-2 text-xs">
    <button class="tbtn px-4 rounded-lg" @click="car.mirror='折叠';toast('已折叠')">折叠</button>
    <button class="tbtn px-4 rounded-lg" @click="toast('调节模式：↑↓←→')">调节</button></div></div>
  <div class="flex justify-between items-center"><span>倒车影像（模拟）</span>
   <button class="tbtn bg-card2 rounded-xl px-4" @click="car.rearCam=!car.rearCam" :class="car.rearCam?'text-accent':''">{{car.rearCam?'开启':'关闭'}}</button></div>
 </div>
 <div v-if="car.rearCam" class="bg-black rounded-2xl border border-accent p-4 text-center">
  <div class="text-lg">🎥 倒车影像（模拟）</div>
  <div class="text-xs text-muted mt-1">广角模拟 · 挂R联动雷达</div>
  <div class="grid grid-cols-3 gap-2 mt-3 text-xs text-warn">
   <div class="bg-white/5 rounded p-2">左 1.5m</div><div class="bg-white/5 rounded p-2">中 0.8m</div><div class="bg-white/5 rounded p-2">右 1.3m</div></div>
 </div>
</div>`,
 setup(){car.rearCam=false;return{car}}};

const Cam360={template:`
<div class="p-4 space-y-3 pop">
 <div class="bg-card rounded-2xl border border-line p-5">
  <div class="text-lg font-bold mb-3">📷 360° 全景影像（模拟）</div>
  <div class="relative bg-black/50 rounded-xl h-56 flex items-center justify-center overflow-hidden">
   <div class="absolute inset-5 border-2 border-dashed border-accent/30 rounded-lg"></div>
   <div class="text-5xl">🚙</div>
   <div class="absolute top-2 left-3 text-xs text-accent">前 · 广角</div>
   <div class="absolute top-2 right-3 text-xs text-muted">俯视合成</div>
  </div>
  <div class="grid grid-cols-4 gap-2 mt-3 text-xs">
   <button class="tbtn bg-card2 rounded-xl" @click="toast('视角: 前')">前</button>
   <button class="tbtn bg-card2 rounded-xl" @click="toast('视角: 后')">后</button>
   <button class="tbtn bg-card2 rounded-xl" @click="toast('视角: 左')">左</button>
   <button class="tbtn bg-card2 rounded-xl" @click="toast('视角: 右')">右</button>
  </div>
 </div>
</div>`};

const Park={template:`
<div class="p-4 space-y-3 pop">
 <div class="bg-card rounded-2xl border border-line p-5 text-center">
  <div class="text-lg font-bold mb-3">🅿️ 泊车辅助</div>
  <div class="bg-black/50 rounded-xl h-40 flex items-center justify-center relative">
   <div class="absolute inset-x-4 top-3 border-t-2 border-dashed border-warn/50"></div>
   <div class="absolute inset-x-4 bottom-3 border-t-2 border-dashed border-warn/50"></div>
   <div class="text-3xl font-bold" :class="car.radarDist&&car.radarDist<0.5?'text-danger':'text-ok'">
    {{car.radarDist?car.radarDist+' m':'挂 R 激活雷达'}}</div>
  </div>
  <div class="grid grid-cols-2 gap-3 mt-3">
   <button class="tbtn bg-accent rounded-xl font-bold text-slate-900" @click="setGear('R')">🅿️ 模拟倒车</button>
   <button class="tbtn bg-card2 rounded-xl" @click="setGear('P')">P 挡</button>
  </div>
 </div>
</div>`};

const Settings={template:`
<div class="p-4 space-y-3 pop">
 <div class="bg-card rounded-2xl border border-line p-4 space-y-4 text-sm">
  <div class="flex justify-between items-center"><span>屏幕亮度 {{car.bright}}%</span>
   <input type="range" min="20" max="100" v-model.number="car.bright" class="w-40"></div>
  <div class="flex justify-between"><span>座舱主题（深/浅）</span>
   <button class="tbtn px-4 rounded-xl font-bold" :class="car.dark?'bg-accent text-slate-900':'bg-card2'" @click="car.dark=!car.dark">{{car.dark?'深色':'浅色'}}</button></div>
  <div class="flex justify-between items-center"><span>WiFi</span>
   <button class="tbtn px-4 rounded-xl" :class="car.wifi?'bg-ok text-slate-900':'bg-card2'" @click="car.wifi=!car.wifi">{{car.wifi?'已连接':'未连接'}}</button></div>
  <div class="flex justify-between"><span>蓝牙 · {{car.bt}}</span><span class="text-ok">已连接</span></div>
  <div class="flex justify-between cursor-pointer" @click="toast('OTA模拟: 已是最新 '+car.ota)"><span>OTA 更新</span><span class="text-accent">检查 ›</span></div>
  <div class="flex justify-between text-muted"><span>关于本机</span><span>RX5 HMI {{car.ota}}</span></div>
 </div>
</div>`};

createApp({
 components:{Dock},data(){return{car}},
 computed:{pages(){return{home:Home,nav:NavPage,air:AirPage,media:Media,vehicle:Vehicle,cam:Cam360,park:Park,settings:Settings}},
  page(){return this.pages[this.car.page]||Home}},
 methods:{setGear,startVoice,sendCmd},
 template:`
<div class="h-full flex flex-col" :class="!car.dark?'bg-slate-100 text-slate-900':''">
 <div class="h-11 flex justify-between items-center px-5 text-sm opacity-80 shrink-0 bg-card">
  <div><b class="text-base">{{car.speed}}</b> km/h · 续航 {{car.range}}km</div>
  <div>{{car.weather}}</div><div>{{car.net}} · {{car.clock}}</div>
 </div>
 <div class="flex-1 relative overflow-hidden">
  <component :is="page" :key="car.page" class="absolute inset-0 overflow-y-auto"/>
  <div v-if="car.nav&&car.page!=='nav'" class="absolute top-3 left-3 z-20 bg-card border border-accent rounded-2xl px-4 py-2 text-sm slide">
   🧭 <b class="text-accent">↑</b> 300m 右转 · 12分钟</div>
  <div v-if="car.voice" class="absolute left-4 right-4 bottom-24 z-40 bg-card border border-line rounded-2xl p-4 text-center pop">
   <div class="w-12 h-12 mx-auto rounded-full bg-accent text-2xl flex items-center justify-center animate-pulse">🎙</div>
   <div class="text-sm mt-2 opacity-80">{{car.voiceTxt}}</div></div>
  <div v-if="car.gear==='D'" class="absolute inset-0 bg-black/85 z-30 flex flex-col items-center justify-center gap-3">
   <div class="text-5xl">🚗</div><div class="text-lg font-bold">行驶中（D 挡）</div>
   <div class="text-sm opacity-60">视频与复杂娱乐已暂停 · 请安心驾驶</div>
   <button class="tbtn mt-2 border border-accent text-accent rounded-full px-6" @click="car.gear='P';toast('已模拟切回P挡(演示)')">演示切回 P 挡</button>
  </div>
  <div id="toast" class="absolute bottom-24 left-1/2 -translate-x-1/2 bg-card border border-line px-5 py-2 rounded-full text-sm transition-opacity opacity-0 z-50"></div>
 </div>
 <Dock :p="car.page" @go="k=>{car.page=k;if(k==='voice')startVoice()}"/>
</div>`
}).mount('#app');
console.log('[RX5 HMI] mounted · 页面切换/主控逻辑就绪');
