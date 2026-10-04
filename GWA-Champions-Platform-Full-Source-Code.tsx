import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import QRCode from "qrcode";

/* ═══════ TYPES ═══════ */
type Role = "teacher" | "student";
type View = "home"|"showcase"|"classBoard"|"allWorks"|"winners"|"adminLogin"|"admin"|"teacherLogin"|"teacher"|"studentLogin"|"student";
type User = { id:string; fullName:string; role:Role; grade:string; section:string; teacherId?:string; username:string; password:string; createdAt:string; active:boolean };
type Comp = { id:string; nameAr:string; nameEn:string; logoUrl:string; grades:string[]; rulesAr:string; rulesEn:string; closingDate:string; certDate:string; started:boolean; archived:boolean; criteria:string[]; evalCode:string };
type Work = { id:string; competitionId:string; studentId:string; title:string; desc:string; videoUrl:string; thumb:string; at:string };
type Ev = { id:string; workId:string; scores:number[]; at:string; approved?:boolean };
type Like = { workId:string; type:"friend"|"parent"|"teacher"; name:string; at:string };
type Cert = { id:string; workId:string; studentId:string; compId:string; compName:string; score:number; rank:number; grade:string; section:string; issuedAt:string };
type ClassInfo = { grade:string; section:string; teacherId:string; teacherName:string; count:number };
type ArchiveEntry = { id:string; compId:string; compName:string; compNameAr:string; archivedAt:string; winners:Array<{rank:number; studentName:string; grade:string; section:string; score:number}> };

/* ═══════ CONSTANTS ═══════ */
const SK = { u:"gwa-u-v8", c:"gwa-c-v8", w:"gwa-w-v8", e:"gwa-e-v8", l:"gwa-l-v8", cert:"gwa-cert-v8", arch:"gwa-arch-v8", pw:"gwa-pw-v8" };
const GRADES = ["Grade 1","Grade 2","Grade 3","Grade 4","Grade 5","Grade 6","Grade 7","Grade 8","Grade 9","Grade 10","Grade 11","Grade 12"];
const SECTIONS = ["1A","1B","2A","2B","3A","3B","4A","4B","5A","5B","6X","6Y","7A","7B","8A","8B"];
const DEFAULT_CRITERIA = ["Creativity • الإبداع","Clarity • الوضوح","Content • المحتوى","Language • اللغة","Confidence • الثقة","Design • التصميم","Compliance • الالتزام","Organization • التنظيم","Technology • التقنية","Final Impact • الإخراج"];
const MEDALS = ["🥇","🥈","🥉","4️⃣","5️⃣","6️⃣","7️⃣","8️⃣","9️⃣","🔟"];

/* ═══════ COMPETITION PRESETS WITH SPECIFIC CRITERIA ═══════ */
type CompPreset = { nameEn:string; nameAr:string; criteria:string[] };
const COMP_PRESETS: CompPreset[] = [
  { nameEn:"Quran Recitation Competition", nameAr:"مسابقة تلاوة القرآن الكريم", criteria:[
    "Tajweed Rules • أحكام التجويد","Letter Pronunciation • مخارج الحروف","Voice Beauty • جمال الصوت",
    "Pausing & Starting • الوقف والابتداء","Emotional Expression • تصوير المعنى","Breath Control • التحكم بالنفس",
    "Memorization Accuracy • دقة الحفظ","Rhythm & Flow • الإيقاع والانسيابية","Confidence • الثقة بالنفس","Overall Impression • الانطباع العام"
  ]},
  { nameEn:"Arabic Public Speaking", nameAr:"مسابقة التحدث بالعربية الفصحى", criteria:[
    "Pronunciation • سلامة النطق","Grammar Accuracy • السلامة النحوية","Vocabulary Richness • ثراء المفردات",
    "Fluency • الطلاقة","Body Language • لغة الجسد","Eye Contact • التواصل البصري",
    "Content Depth • عمق المحتوى","Audience Engagement • تفاعل الجمهور","Time Management • إدارة الوقت","Persuasion • القدرة على الإقناع"
  ]},
  { nameEn:"Arabic Creative Writing", nameAr:"مسابقة الكتابة الإبداعية", criteria:[
    "Original Idea • أصالة الفكرة","Language Style • أسلوب اللغة","Grammar & Spelling • النحو والإملاء",
    "Story Structure • بنية القصة","Character Development • تطور الشخصيات","Imagery & Description • التصوير والوصف",
    "Emotional Impact • التأثير العاطفي","Vocabulary Range • تنوع المفردات","Creativity • الإبداع","Conclusion • قوة الخاتمة"
  ]},
  { nameEn:"Arabic Poetry Recitation", nameAr:"مسابقة إلقاء الشعر العربي", criteria:[
    "Pronunciation • سلامة النطق","Rhythm & Meter • الوزن والإيقاع","Emotional Delivery • الأداء العاطفي",
    "Voice Modulation • تنويع الصوت","Memorization • الحفظ","Stage Presence • الحضور المسرحي",
    "Understanding of Meaning • فهم المعنى","Audience Connection • التواصل مع الجمهور","Gestures • الإيماءات","Overall Performance • الأداء الكلي"
  ]},
  { nameEn:"Arabic Reading Competition", nameAr:"مسابقة القراءة العربية", criteria:[
    "Reading Fluency • طلاقة القراءة","Comprehension • الفهم والاستيعاب","Critical Analysis • التحليل النقدي",
    "Summary Skills • مهارات التلخيص","Vocabulary Usage • استخدام المفردات","Presentation Quality • جودة العرض",
    "Book Selection • اختيار الكتاب","Personal Reflection • التأمل الشخصي","Research Depth • عمق البحث","Creative Expression • التعبير الإبداعي"
  ]},
  { nameEn:"Arabic Calligraphy", nameAr:"مسابقة الخط العربي", criteria:[
    "Letter Form • شكل الحروف","Consistency • التناسق","Spacing • المسافات",
    "Line Quality • جودة الخط","Artistic Composition • التكوين الفني","Creativity • الإبداع",
    "Color Usage • استخدام الألوان","Neatness • النظافة","Adherence to Rules • الالتزام بالقواعد","Overall Aesthetics • الجمال العام"
  ]},
  { nameEn:"Islamic Knowledge Quiz", nameAr:"مسابقة المعارف الإسلامية", criteria:[
    "Quran Knowledge • معرفة القرآن","Hadith Knowledge • معرفة الحديث","Fiqh Understanding • فهم الفقه",
    "Seerah Knowledge • معرفة السيرة","Islamic History • التاريخ الإسلامي","Speed of Answer • سرعة الإجابة",
    "Accuracy • الدقة","Explanation Quality • جودة التفسير","Confidence • الثقة","Sportsmanship • الروح الرياضية"
  ]},
  { nameEn:"Arabic Spelling Bee", nameAr:"مسابقة الإملاء العربي", criteria:[
    "Spelling Accuracy • دقة الإملاء","Dictation Speed • سرعة الإملاء","Hamza Rules • قواعد الهمزة",
    "Ta Marbuta/Ha • التاء المربوطة والهاء","Alef Rules • قواعد الألف","Punctuation • علامات الترقيم",
    "Word Derivation • اشتقاق الكلمات","Handwriting • جودة الخط","Listening Skills • مهارات الاستماع","Overall Performance • الأداء الكلي"
  ]},
  { nameEn:"Social Studies Project", nameAr:"مشروع الدراسات الاجتماعية", criteria:[
    "Research Quality • جودة البحث","Topic Relevance • ملاءمة الموضوع","Data Presentation • عرض البيانات",
    "Critical Thinking • التفكير النقدي","Creativity • الإبداع","Visual Design • التصميم البصري",
    "Historical Accuracy • الدقة التاريخية","Oral Presentation • العرض الشفهي","Teamwork • العمل الجماعي","Conclusions • الاستنتاجات"
  ]},
  { nameEn:"Arabic Debate Competition", nameAr:"مسابقة المناظرة بالعربية", criteria:[
    "Argument Strength • قوة الحجة","Rebuttal Skills • مهارات الرد","Evidence Usage • استخدام الأدلة",
    "Language Proficiency • إتقان اللغة","Persuasion • الإقناع","Logical Structure • البنية المنطقية",
    "Delivery Style • أسلوب الإلقاء","Time Management • إدارة الوقت","Teamwork • العمل الجماعي","Overall Impact • التأثير العام"
  ]},
];

function getClassColor(grade: string) {
  const colors = [
    { bg:"rgba(22,151,232,0.35)", border:"#1697E8", glow:"rgba(22,151,232,0.15)" },
    { bg:"rgba(253,181,31,0.35)", border:"#FDB51F", glow:"rgba(253,181,31,0.15)" },
    { bg:"rgba(16,176,155,0.35)", border:"#10B09B", glow:"rgba(16,176,155,0.15)" },
    { bg:"rgba(226,20,85,0.35)", border:"#E21455", glow:"rgba(226,20,85,0.15)" },
    { bg:"rgba(139,92,246,0.35)", border:"#8B5CF6", glow:"rgba(139,92,246,0.15)" },
    { bg:"rgba(236,72,153,0.35)", border:"#EC4899", glow:"rgba(236,72,153,0.15)" },
  ];
  let hash = 0;
  for (let i = 0; i < grade.length; i++) hash = grade.charCodeAt(i) + ((hash << 5) - hash);
  return colors[Math.abs(hash) % colors.length];
}

/* ═══════ SEED ═══════ */
const T=["https://images.pexels.com/photos/39701530/pexels-photo-39701530.jpeg?auto=compress&cs=tinysrgb&fit=crop&h=627&w=1200","https://images.pexels.com/photos/33750570/pexels-photo-33750570.jpeg?auto=compress&cs=tinysrgb&fit=crop&h=627&w=1200","https://images.pexels.com/photos/33750571/pexels-photo-33750571.jpeg?auto=compress&cs=tinysrgb&fit=crop&h=627&w=1200","https://images.pexels.com/photos/20627703/pexels-photo-20627703.jpeg?auto=compress&cs=tinysrgb&fit=crop&h=627&w=1200","https://images.pexels.com/photos/8466718/pexels-photo-8466718.jpeg?auto=compress&cs=tinysrgb&fit=crop&h=627&w=1200","https://images.pexels.com/photos/6209557/pexels-photo-6209557.jpeg?auto=compress&cs=tinysrgb&fit=crop&h=627&w=1200","https://images.pexels.com/photos/18800463/pexels-photo-18800463.jpeg?auto=compress&cs=tinysrgb&fit=crop&h=627&w=1200","https://images.pexels.com/photos/32229325/pexels-photo-32229325.jpeg?auto=compress&cs=tinysrgb&fit=crop&h=627&w=1200","https://images.unsplash.com/photo-1522202176988-66273c2fd55f?auto=format&fit=crop&w=800&q=80","https://images.unsplash.com/photo-1516321318423-f06f85e504b3?auto=format&fit=crop&w=800&q=80","https://images.unsplash.com/photo-1532012197267-da84d127e765?auto=format&fit=crop&w=800&q=80","https://images.unsplash.com/photo-1503676260728-1c00da094a0b?auto=format&fit=crop&w=800&q=80"];
const V="https://www.w3schools.com/html/mov_bbb.mp4";
const qNames=["Yusuf Al-Rashid","Amina Tariq","Hassan Omar","Fatima Noor","Khalid Ibrahim","Zahra Saeed","Abdullah Hamad","Maryam Faisal","Sami Waleed","Nadia Jamal","Ali Mustafa","Reem Salah"];
const tNames=["Emma Johnson","Lucas Chen","Sophie Martin","James Wilson","Mia Rodriguez","Leo Patel","Ava Kim","Noah Müller","Lily Tanaka","Oscar Brown","Clara Ivanova","Max Torres"];
const pNames=["أحمد بن خالد","فاطمة الزهراء","محمد الحربي","نورة السالم","عبدالرحمن العتيبي","هند المنصور","سعود الدوسري","ريم الشمري","خالد النعيمي","لينا العبدالله","ماجد الهاشمي","سارة القحطاني"];
const rNames=["يوسف المطيري","آمنة الكتبي","حمدان السويدي","مريم البلوشي","عمر الشامسي","هالة المزروعي","سالم الظاهري","نورا الفلاسي","فيصل الكعبي","دانة المهيري","راشد النقبي","شيخة الحبسي"];
const uaeNames=["سلطان المنصوري","موزة الكتبي","حمد الشحي","عليا الظاهري","سعيد النيادي","شما البلوشي","ماجد الرميثي","هند الزعابي","أحمد المزروعي","فاطمة الحمادي","خليفة الشامسي","مريم السويدي"];

// Seed students (60 demo students across competitions)
const demoStudents:User[]=[];
[...qNames,...tNames,...pNames,...rNames,...uaeNames].forEach((n,i)=>{
  const grade=`Grade ${4+Math.floor(i%9)}`;const sec=`${4+Math.floor(i%9)}${["A","B","X"][i%3]}`;
  demoStudents.push({id:`ds${i}`,fullName:n,role:"student",grade,section:sec,teacherId:i%2===0?"t1":"t2",username:`demo${i}`,password:"GWA2050",createdAt:"",active:true});
});

const seedU: User[] = [
  { id:"t1", fullName:"Ms Heba", role:"teacher", grade:"Grade 6", section:"6X", username:"heba.6x", password:"Teach2050", createdAt:"", active:true },
  { id:"t2", fullName:"Mr Ahmed", role:"teacher", grade:"Grade 5", section:"5A", username:"ahmed.5a", password:"Teach2050", createdAt:"", active:true },
  { id:"s1", fullName:"Mariam Salem", role:"student", grade:"Grade 6", section:"6X", teacherId:"t1", username:"mariam.salem", password:"GWA2050", createdAt:"", active:true },
  { id:"s2", fullName:"Omar Khalid", role:"student", grade:"Grade 6", section:"6X", teacherId:"t1", username:"omar.khalid", password:"GWA2050", createdAt:"", active:true },
  { id:"s3", fullName:"Sara Mansour", role:"student", grade:"Grade 5", section:"5A", teacherId:"t2", username:"sara.mansour", password:"GWA2050", createdAt:"", active:true },
  ...demoStudents,
];

// ── 5 SEEDED COMPETITIONS ──
const seedC:Comp[]=[
  {id:"cq1",nameEn:"Holy Quran Recitation Competition",nameAr:"مسابقة تلاوة القرآن الكريم",logoUrl:"/gems-logo.png",grades:["Grade 4","Grade 5","Grade 6","Grade 7","Grade 8"],rulesAr:"١- يتلو الطالب آيات من القرآن الكريم\n٢- يُقيّم على أساس التجويد والأداء\n٣- مدة الفيديو ٣ دقائق\n٤- الالتزام بأحكام التلاوة",rulesEn:"1- Student recites verses from the Holy Quran\n2- Evaluated on Tajweed and delivery\n3- Video duration: 3 minutes\n4- Adherence to recitation rules",closingDate:"2027-02-01",certDate:"2027-02-15",started:true,archived:false,criteria:COMP_PRESETS[0].criteria,evalCode:"QURAN1"},
  {id:"ct1",nameEn:"Arabic Big Talk",nameAr:"مسابقة التحدث بالعربية",logoUrl:"/gems-logo.png",grades:["Grade 4","Grade 5","Grade 6","Grade 7","Grade 8"],rulesAr:"١- يتحدث الطالب بالعربية الفصحى لمدة ٢-٣ دقائق\n٢- يُقيّم على الطلاقة والنطق ولغة الجسد\n٣- الموضوع حر\n٤- مفتوحة للطلاب العرب وغير العرب",rulesEn:"1- Student speaks in Arabic for 2-3 minutes\n2- Evaluated on fluency, pronunciation & body language\n3- Free topic choice\n4- Open to Arab and non-Arab students",closingDate:"2027-02-01",certDate:"2027-02-15",started:true,archived:false,criteria:COMP_PRESETS[1].criteria,evalCode:"TALK1"},
  {id:"cp1",nameEn:"Classical Arabic Poetry Recitation",nameAr:"مسابقة الشعر العربي الأصيل",logoUrl:"/gems-logo.png",grades:["Grade 5","Grade 6","Grade 7","Grade 8","Grade 9"],rulesAr:"١- يُلقي الطالب قصيدة من الشعر العربي الكلاسيكي\n٢- التقييم على الأداء والإيقاع والتعبير\n٣- مدة الإلقاء ٢-٤ دقائق\n٤- يُفضل الحفظ الكامل",rulesEn:"1- Student recites a classical Arabic poem\n2- Evaluated on delivery, rhythm & expression\n3- Duration: 2-4 minutes\n4- Full memorization preferred",closingDate:"2027-02-01",certDate:"2027-02-15",started:true,archived:false,criteria:COMP_PRESETS[3].criteria,evalCode:"POETRY1"},
  {id:"cr1",nameEn:"Arabic Loud Reading Competition",nameAr:"مسابقة القراءة الجهرية",logoUrl:"/gems-logo.png",grades:["Grade 3","Grade 4","Grade 5","Grade 6","Grade 7"],rulesAr:"١- يقرأ الطالب نصاً عربياً قراءة جهرية واضحة\n٢- التقييم على النطق والطلاقة والفهم\n٣- مدة القراءة ٢-٣ دقائق\n٤- النص يُختار من الأدب العربي",rulesEn:"1- Student reads an Arabic text aloud clearly\n2- Evaluated on pronunciation, fluency & comprehension\n3- Duration: 2-3 minutes\n4- Text selected from Arabic literature",closingDate:"2027-02-01",certDate:"2027-02-15",started:true,archived:false,criteria:COMP_PRESETS[4].criteria,evalCode:"READ1"},
  {id:"cu1",nameEn:"UAE: Land of Goodness",nameAr:"مسابقة الاجتماعيات: الإمارات وطن الخير",logoUrl:"/gems-logo.png",grades:["Grade 4","Grade 5","Grade 6","Grade 7","Grade 8"],rulesAr:"١- يقدم الطالب مشروعاً عن الإمارات العربية المتحدة\n٢- يمكن أن يكون فيديو أو عرض رقمي\n٣- يركز على التاريخ أو الجغرافيا أو القيم الوطنية\n٤- الأصالة والإبداع أساس التقييم",rulesEn:"1- Student presents a project about the UAE\n2- Can be a video or digital presentation\n3- Focus on history, geography or national values\n4- Originality and creativity are key",closingDate:"2027-02-01",certDate:"2027-02-15",started:true,archived:false,criteria:COMP_PRESETS[8].criteria,evalCode:"UAE1"},
];

// ── 60 SEEDED WORKS (12 per competition) ──
const seedW:Work[]=[];
const qTitles=["Surah Al-Fatiha Recitation","Surah Ar-Rahman Beautiful","Surah Al-Mulk Melodious","Surah Yasin Heartfelt","Surah Al-Kahf Peaceful","Surah Maryam Emotional","Juz Amma Complete","Surah Al-Baqarah Excerpt","Surah Al-Imran Selection","Surah An-Noor Recitation","Surah Ad-Duha Beautiful","Surah Al-Inshirah Soothing"];
const tTitles=["My Journey Learning Arabic","Why I Love Arabic Language","Arabic Changed My Life","Speaking Arabic at School","My Arabic Teacher Inspired Me","Arabic: Bridge Between Cultures","Discovering Arabic Poetry","Arabic Words I Love Most","My First Arabic Conversation","Arabic in Daily Life","The Beauty of Arabic Script","Arabic: Language of the Heart"];
const pTitles=["قصيدة المتنبي - على قدر أهل العزم","قصيدة أبو فراس - أراك عصي الدمع","قصيدة نزار قباني - أحبك","شعر أحمد شوقي - قم للمعلم","قصيدة إيليا أبو ماضي - الطلاسم","قصيدة الإمام الشافعي - دع الأيام","أبيات من المعلقات - امرؤ القيس","شعر جرير - لولا الحياء","قصيدة حافظ إبراهيم - اللغة العربية","شعر أبو العتاهية - الزهد","قصيدة البحتري - وصف الربيع","شعر أبو تمام - السيف أصدق"];
const rTitles=["قراءة من كتاب الأيام - طه حسين","قراءة من رواية أولاد حارتنا","قراءة نص الجاحظ - البخلاء","قراءة من كتاب كليلة ودمنة","قراءة نص أدبي - جبران خليل جبران","قراءة من مقامات الحريري","قراءة نثرية - مصطفى لطفي المنفلوطي","قراءة من ألف ليلة وليلة","قراءة نص طه حسين - على هامش السيرة","قراءة إبداعية - ميخائيل نعيمة","قراءة من رسالة الغفران","قراءة من حي بن يقظان"];
const uTitles=["UAE: 7 Emirates United","Expo Dubai: Window to Future","Sheikh Zayed: Father of Nation","UAE Mars Mission: Hope Probe","Dubai: City of Innovation","Abu Dhabi: Capital of Culture","Emirati Heritage & Traditions","UAE National Day Celebration","Masdar City: Green Future","Louvre Abu Dhabi: Art Bridge","UAE Space Program Vision","Sharjah: Cultural Capital"];

qTitles.forEach((t,i)=>seedW.push({id:`wq${i}`,competitionId:"cq1",studentId:`ds${i}`,title:t,desc:"",videoUrl:V,thumb:T[i%T.length],at:new Date(Date.now()-i*86400000).toISOString()}));
tTitles.forEach((t,i)=>seedW.push({id:`wt${i}`,competitionId:"ct1",studentId:`ds${12+i}`,title:t,desc:"",videoUrl:V,thumb:T[(i+4)%T.length],at:new Date(Date.now()-i*86400000).toISOString()}));
pTitles.forEach((t,i)=>seedW.push({id:`wp${i}`,competitionId:"cp1",studentId:`ds${24+i}`,title:t,desc:"",videoUrl:V,thumb:T[(i+2)%T.length],at:new Date(Date.now()-i*86400000).toISOString()}));
rTitles.forEach((t,i)=>seedW.push({id:`wr${i}`,competitionId:"cr1",studentId:`ds${36+i}`,title:t,desc:"",videoUrl:V,thumb:T[(i+6)%T.length],at:new Date(Date.now()-i*86400000).toISOString()}));
uTitles.forEach((t,i)=>seedW.push({id:`wu${i}`,competitionId:"cu1",studentId:`ds${48+i}`,title:t,desc:"",videoUrl:V,thumb:T[(i+1)%T.length],at:new Date(Date.now()-i*86400000).toISOString()}));

/* ═══════ HELPERS ═══════ */
function cn(...c:Array<string|false|null|undefined>){return c.filter(Boolean).join(" ")}
function mid(p:string){return `${p}-${Math.random().toString(36).slice(2,8)}`}
function ld<T>(k:string,f:T):T{try{const v=localStorage.getItem(k);return v?JSON.parse(v):f}catch{return f}}
function sv(k:string,v:unknown){localStorage.setItem(k,JSON.stringify(v))}
function tot(s:number[]){return s.reduce((a,b)=>a+b,0)}
function slug(n:string){return n.trim().toLowerCase().replace(/[^a-z0-9\s]/g,"").replace(/\s+/g,".").replace(/^\.+|\.+$/g,"")}
function mkUser(name:string,role:Role,grade:string,section:string,tid:string|undefined,users:User[]):User{
  let u=slug(name)||"gwa.user";let c=1;while(users.some(x=>x.username===u)){u=slug(name)+c;c++}
  const num=String(users.length+2050).slice(-4);
  return{id:mid(role.charAt(0)),fullName:name.trim(),role,grade,section,teacherId:tid,username:u,password:(role==="teacher"?"Teach":"GWA")+num,createdAt:new Date().toISOString(),active:true}
}

/* ═══════ SHARED UI ═══════ */
function Bg(){return<div className="pointer-events-none fixed inset-0 overflow-hidden z-0"><div className="absolute left-[-5%] top-[-5%] h-[500px] w-[500px] rounded-full bg-[#1697E8]/15 blur-[120px]"/><div className="absolute right-[-5%] top-[10%] h-[500px] w-[500px] rounded-full bg-[#FDB51F]/10 blur-[120px]"/><div className="absolute bottom-[-5%] left-[30%] h-[500px] w-[500px] rounded-full bg-[#E21455]/8 blur-[120px]"/><div className="absolute inset-0 opacity-[0.04]" style={{backgroundImage:"radial-gradient(circle,#fff 1px,transparent 1px)",backgroundSize:"32px 32px"}}/></div>}
function Logo({s}:{s?:boolean}){return<div className="rounded-2xl bg-white p-1.5 shadow-xl shadow-black/20 shrink-0"><img src="/gems-logo.png" alt="GEMS" className={cn(s?"h-10":"h-14","w-auto object-contain")}/></div>}
function Pill({children,t="blue"}:{children:React.ReactNode;t?:string}){
  const m:Record<string,string>={blue:"border-sky-400/20 bg-sky-500/15 text-sky-200",gold:"border-[#FDB51F]/25 bg-[#FDB51F]/15 text-[#FFE29A]",green:"border-emerald-400/20 bg-emerald-500/15 text-emerald-200",pink:"border-pink-400/20 bg-pink-500/15 text-pink-200",slate:"border-white/10 bg-white/5 text-white/60"};
  return<span className={cn("inline-flex rounded-full border px-3 py-1 text-xs font-bold",m[t]||m.blue)}>{children}</span>
}
function FL({children}:{children:React.ReactNode}){return<label className="mb-2 block text-sm font-bold text-white/70">{children}</label>}
function Inp(p:React.InputHTMLAttributes<HTMLInputElement>){return<input {...p} className={cn("w-full rounded-2xl border border-white/10 bg-slate-950/35 px-4 py-3 text-white outline-none placeholder:text-white/20 focus:border-[#FDB51F] focus:ring-2 focus:ring-[#FDB51F]/20",p.className)}/>}
function Sel(p:React.SelectHTMLAttributes<HTMLSelectElement>){return<select {...p} className={cn("w-full rounded-2xl border border-white/10 bg-slate-950/35 px-4 py-3 text-white outline-none focus:border-[#FDB51F]",p.className)}/>}
function TA(p:React.TextareaHTMLAttributes<HTMLTextAreaElement>){return<textarea {...p} className={cn("w-full rounded-2xl border border-white/10 bg-slate-950/35 px-4 py-3 text-white outline-none placeholder:text-white/20 focus:border-[#FDB51F] resize-none",p.className)}/>}
function Stat({l,v,s}:{l:string;v:string|number;s:string}){return<div className="rounded-3xl border border-white/10 bg-white/5 p-5 backdrop-blur-sm"><div className="text-[10px] font-bold uppercase tracking-[0.24em] text-white/40">{l}</div><div className="mt-2 text-3xl font-black text-white">{v}</div><div className="mt-1 text-sm text-white/50">{s}</div></div>}

/* ── TEACHER LOGIN QUICK MODAL (for class board / global works) ── */
function TeacherQuickLogin({users,onOk,onClose}:{users:User[];onOk:(t:User)=>void;onClose:()=>void}){
  const[u,setU]=useState(""); const[p,setP]=useState(""); const[err,setErr]=useState("");
  function sub(e:React.FormEvent){e.preventDefault(); const t=users.find(x=>x.role==="teacher"&&x.username===u.trim()&&x.password===p&&x.active); if(t)onOk(t); else setErr("Invalid credentials")}
  return(
    <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-md flex items-center justify-center p-4">
      <div className="w-full max-w-sm rounded-3xl border border-white/10 bg-[#0f1535] p-6 shadow-2xl">
        <div className="flex justify-between items-center mb-6"><h2 className="text-xl font-black text-white">Teacher Login • دخول المعلم</h2><button onClick={onClose} className="text-white/40 hover:text-white font-bold">✕</button></div>
        <form onSubmit={sub} className="space-y-4">
          <div><FL>Username</FL><Inp value={u} onChange={e=>setU(e.target.value)} placeholder="e.g. heba.6x"/></div>
          <div><FL>Password</FL><Inp type="password" value={p} onChange={e=>setP(e.target.value)} placeholder="Password"/></div>
          {err&&<div className="rounded-xl bg-rose-500/10 border border-rose-400/20 px-4 py-3 text-sm text-rose-200 font-bold">{err}</div>}
          <button type="submit" className="w-full rounded-2xl bg-sky-500 text-white font-black py-3.5 hover:bg-sky-400">Sign In ✓</button>
        </form>
      </div>
    </div>
  );
}

/* ── EVALUATION EDITOR MODAL ── */
function EvalModal({workId,evals,setEvals,onClose,teacherName,criteria=DEFAULT_CRITERIA}:{workId:string;evals:Ev[];setEvals:React.Dispatch<React.SetStateAction<Ev[]>>;onClose:()=>void;teacherName:string;criteria:string[]}){
  const ex=evals.find(e=>e.workId===workId); const[scores,setScores]=useState<number[]>(ex?[...ex.scores]:Array(10).fill(5));
  function save(){if(ex){setEvals(p=>p.map(e=>e.workId===workId?{...e,scores,at:new Date().toISOString()}:e))}else{setEvals(p=>[...p,{id:mid("e"),workId,scores,at:new Date().toISOString(),approved:false}])}onClose()}
  return(
    <div className="fixed inset-0 z-50 bg-black/85 backdrop-blur-md flex items-center justify-center p-4">
      <div className="bg-[#0f1535] border border-white/10 rounded-[2rem] p-6 w-full max-w-3xl max-h-[90vh] overflow-y-auto shadow-2xl">
        <div className="flex justify-between items-center mb-6"><div><h2 className="text-2xl font-black text-[#FDB51F]">Evaluation • تحرير التقييم</h2><p className="text-sm text-white/50">Teacher: {teacherName}</p></div><button onClick={onClose} className="rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm font-bold text-white">✕</button></div>
        <div className="space-y-4">{criteria.map((cr:string,i:number)=>(<div key={i} className="bg-white/5 p-4 rounded-2xl border border-white/10"><div className="flex justify-between mb-2"><span className="font-bold text-white/80 text-sm">{cr}</span><span className="font-black text-sky-300">{scores[i]}/10</span></div><input type="range" min="0" max="10" value={scores[i]||0} onChange={e=>{const n=[...scores];n[i]=+e.target.value;setScores(n)}} className="w-full accent-sky-400 cursor-pointer"/></div>))}</div>
        <div className="mt-6 flex flex-col sm:flex-row gap-4 items-center justify-between"><div className="rounded-2xl border border-[#FDB51F]/20 bg-[#FDB51F]/10 px-5 py-4 text-sm font-black text-[#FFE29A]">Total: {tot(scores)}/100</div><button onClick={save} className="rounded-2xl bg-[#FDB51F] px-6 py-4 text-sm font-black text-slate-950 shadow-lg hover:bg-[#F7C84A]">Save ✅</button></div>
      </div>
    </div>
  );
}

/* ── LIKES PANEL ── */
function LikesPanel({workId,likes,setLikes}:{workId:string;likes:Like[];setLikes:React.Dispatch<React.SetStateAction<Like[]>>}){
  const wl=useMemo(()=>likes.filter(l=>l.workId===workId),[likes,workId]);
  const f=wl.filter(l=>l.type==="friend"),p=wl.filter(l=>l.type==="parent"),t=wl.filter(l=>l.type==="teacher");
  function add(type:"friend"|"parent"|"teacher"){setLikes(prev=>[...prev,{workId,type,name:"",at:new Date().toISOString()}])}
  return(
    <div className="mt-4 flex flex-wrap gap-2">
      <button onClick={()=>add("friend")} className="rounded-xl border border-sky-400/20 bg-sky-500/10 px-4 py-2.5 text-xs font-bold text-sky-300 hover:bg-sky-500/20 transition flex-1">👫 Friend ({f.length})</button>
      <button onClick={()=>add("parent")} className="rounded-xl border border-emerald-400/20 bg-emerald-500/10 px-4 py-2.5 text-xs font-bold text-emerald-300 hover:bg-emerald-500/20 transition flex-1">👨‍👩‍👧 Parent ({p.length})</button>
      <button onClick={()=>add("teacher")} className="rounded-xl border border-pink-400/20 bg-pink-500/10 px-4 py-2.5 text-xs font-bold text-pink-300 hover:bg-pink-500/20 transition flex-1">👩‍🏫 Teacher ({t.length})</button>
    </div>
  );
}

/* ── WORK CARD (shared between class board and global works) ── */
function WorkCard({w,users,evals,comps,likes,setLikes,setEvals,setWorks}:{w:Work;users:User[];evals:Ev[];comps:Comp[];likes:Like[];setLikes:React.Dispatch<React.SetStateAction<Like[]>>;setEvals:React.Dispatch<React.SetStateAction<Ev[]>>;setWorks:React.Dispatch<React.SetStateAction<Work[]>>}){
  const stu=users.find(s=>s.id===w.studentId);
  const comp=comps.find(c=>c.id===w.competitionId);
  const ev=evals.find(e=>e.workId===w.id);
  const[viewEval,setViewEval]=useState(false);
  const[editEval,setEditEval]=useState(false);
  const[showCodePrompt,setShowCodePrompt]=useState(false);
  const[codeMode,setCodeMode]=useState<"eval"|"delete">("eval");
  const[codeInput,setCodeInput]=useState("");
  const[codeErr,setCodeErr]=useState("");
  if(!stu)return null;

  function verifyCode(){
    if(!comp){setCodeErr("No competition found");return}
    if(codeInput.trim().toUpperCase()===comp.evalCode?.toUpperCase()){
      setShowCodePrompt(false);
      setCodeInput("");
      setCodeErr("");
      if(codeMode==="eval"){
        setEditEval(true);
      }else{
        setWorks(prev=>prev.filter(item=>item.id!==w.id));
        setEvals(prev=>prev.filter(item=>item.workId!==w.id));
        setLikes(prev=>prev.filter(item=>item.workId!==w.id));
      }
    }else{
      setCodeErr("Invalid evaluation code • كود التصحيح غير صحيح");
    }
  }

  const critList=comp?.criteria?.length?comp.criteria:DEFAULT_CRITERIA;

  return(
    <article className="group rounded-[2rem] border border-white/10 bg-white/5 backdrop-blur-sm overflow-hidden flex flex-col hover:-translate-y-0.5 transition-all">
      <div className="relative bg-black"><video src={w.videoUrl} poster={w.thumb} controls preload="metadata" className="w-full aspect-video object-cover"/></div>
      <div className="p-5 flex-1 flex flex-col">
        {comp&&<div className="text-[10px] font-black uppercase tracking-[0.22em] text-[#1697E8] mb-2">{comp.nameEn}</div>}
        <h3 className="text-lg font-black text-white leading-snug">{w.title}</h3>
        {w.desc&&<p className="mt-2 text-sm text-white/50 leading-relaxed line-clamp-2">{w.desc}</p>}
        <div className="mt-4 pt-4 border-t border-white/10 flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center rounded-full bg-[#1F2E76] text-sm font-black text-[#FDB51F] shrink-0">{stu.fullName[0]}</div>
          <div><p className="text-sm font-black text-white">{stu.fullName}</p><p className="text-xs text-white/40">{stu.grade} • Section {stu.section}</p></div>
        </div>
        <div className="mt-auto pt-4 flex flex-col gap-2">
          {ev?(
            <button onClick={()=>setViewEval(true)} className="w-full rounded-2xl bg-emerald-500/15 border border-emerald-400/20 px-4 py-3 text-sm font-black text-emerald-200 hover:bg-emerald-500/25 transition">
              📊 View Score ({tot(ev.scores)}/100) • مشاهدة التقييم
            </button>
          ):(
            <button onClick={()=>setShowCodePrompt(true)} className="w-full rounded-2xl bg-white/5 border border-white/10 px-4 py-3 text-sm font-bold text-white/40 hover:text-[#FDB51F] hover:border-[#FDB51F]/30 hover:bg-[#FDB51F]/10 transition cursor-pointer text-center">
              ⏳ Under Review • بانتظار التقييم — Click to Evaluate
            </button>
          )}
          {/* Edit existing eval */}
          {ev&&(
            <button onClick={()=>{setCodeMode("eval");setShowCodePrompt(true)}} className="w-full rounded-2xl bg-white/5 border border-white/10 px-4 py-2.5 text-xs font-bold text-white/30 hover:text-[#FDB51F] hover:border-[#FDB51F]/30 transition cursor-pointer text-center">
              ✏️ Edit Evaluation • تعديل التقييم
            </button>
          )}
          <button onClick={()=>{setCodeMode("delete");setShowCodePrompt(true)}} className="w-full rounded-2xl bg-rose-500/10 border border-rose-400/20 px-4 py-2.5 text-xs font-black text-rose-200 hover:bg-rose-500/20 transition cursor-pointer text-center">
            🗑 Delete Video • حذف الفيديو
          </button>
        </div>
        <LikesPanel workId={w.id} likes={likes} setLikes={setLikes}/>
      </div>

      {/* Eval Code Prompt Modal */}
      {showCodePrompt&&(
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 p-4 backdrop-blur-sm" onClick={()=>{setShowCodePrompt(false);setCodeErr("")}}>
          <div className="w-full max-w-sm rounded-3xl border border-white/10 bg-[#0f1535] p-6 shadow-2xl" onClick={e=>e.stopPropagation()}>
            <div className="text-center mb-5">
              <div className="text-3xl mb-2">🔐</div>
              <h3 className="text-xl font-black text-white">Evaluation Code</h3>
              <p className="text-sm text-white/50 mt-1">أدخل كود التصحيح الخاص بالمسابقة</p>
              <p className="text-sm text-white/40 mt-1">Enter the evaluation code provided by admin</p>
            </div>
            <div className="space-y-4">
              <input value={codeInput} onChange={e=>{setCodeInput(e.target.value);if(codeErr)setCodeErr("")}} onKeyDown={e=>e.key==="Enter"&&verifyCode()} placeholder="Evaluation Code" className="w-full rounded-2xl border border-white/10 bg-slate-950/35 px-4 py-3 text-white text-center text-lg font-black tracking-[0.3em] uppercase outline-none placeholder:text-white/20 placeholder:tracking-normal placeholder:font-normal focus:border-[#FDB51F]" autoFocus/>
              {codeErr&&<div className="rounded-xl bg-rose-500/10 border border-rose-400/20 px-4 py-3 text-sm text-rose-200 font-bold text-center">{codeErr}</div>}
              <div className="flex gap-3">
                <button onClick={()=>{setShowCodePrompt(false);setCodeErr("")}} className="flex-1 rounded-2xl border border-white/10 bg-white/5 py-3 text-sm font-bold text-white/60">Cancel</button>
                <button onClick={verifyCode} className="flex-1 rounded-2xl bg-[#FDB51F] py-3 text-sm font-black text-slate-950 hover:bg-[#F7C84A]">Verify & Evaluate ✓</button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Eval view modal */}
      {viewEval&&ev&&(
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 p-4 backdrop-blur-sm" onClick={()=>setViewEval(false)}>
          <div className="w-full max-w-2xl rounded-[2rem] border border-white/10 bg-[#0f1535] p-6 shadow-2xl max-h-[90vh] overflow-y-auto" onClick={e=>e.stopPropagation()}>
            <div className="flex items-center justify-between mb-6"><div><h3 className="text-2xl font-black text-white">Evaluation Scorecard</h3><p className="text-sm text-white/50">لوحة التقييم</p></div><button onClick={()=>setViewEval(false)} className="rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm font-bold text-white hover:bg-white/10">✕</button></div>
            <div className="space-y-3">{critList.map((cr:string,i:number)=>(<div key={i} className="flex items-center justify-between rounded-2xl border border-white/10 bg-white/5 px-4 py-3"><span className="text-sm font-bold text-white/75">{cr}</span><span className="text-lg font-black text-sky-300">{ev.scores[i]}/10</span></div>))}</div>
            <div className="mt-6 flex items-center justify-between rounded-2xl border border-[#FDB51F]/20 bg-[#FDB51F]/10 px-5 py-4"><span className="text-sm font-bold uppercase tracking-[0.24em] text-[#FFE29A]">Final Score</span><span className="text-3xl font-black text-white">{tot(ev.scores)}/100</span></div>
          </div>
        </div>
      )}
      {editEval&&<EvalModal workId={w.id} evals={evals} setEvals={setEvals} onClose={()=>setEditEval(false)} teacherName="Evaluator" criteria={critList}/>}
    </article>
  );
}

/* ════════════════════════════════════
   MAIN APP
   ════════════════════════════════════ */
export default function App(){
  const[view,setView]=useState<View>("home");
  const[users,setUsers]=useState<User[]>(()=>ld(SK.u,seedU));
  const[comps,setComps]=useState<Comp[]>(()=>ld(SK.c,seedC));
  const[works,setWorks]=useState<Work[]>(()=>ld(SK.w,seedW));
  const[evals,setEvals]=useState<Ev[]>(()=>ld(SK.e,[]));
  const[likes,setLikes]=useState<Like[]>(()=>ld(SK.l,[]));
  const[certs,setCerts]=useState<Cert[]>(()=>ld(SK.cert,[]));
  const[archive,setArchive]=useState<ArchiveEntry[]>(()=>ld(SK.arch,[]));
  const[adminPw,setAdminPw]=useState<string>(()=>ld(SK.pw,"Gems@2050"));
  const[selClass,setSelClass]=useState<ClassInfo|null>(null);
  const[selCompId,setSelCompId]=useState<string|null>(null);
  const[teacherSess,setTeacherSess]=useState<User|null>(null);
  const[studentSess,setStudentSess]=useState<User|null>(null);

  useEffect(()=>{sv(SK.u,users)},[users]);
  useEffect(()=>{sv(SK.pw,adminPw)},[adminPw]);
  useEffect(()=>{sv(SK.c,comps)},[comps]);
  useEffect(()=>{sv(SK.w,works)},[works]);
  useEffect(()=>{sv(SK.e,evals)},[evals]);
  useEffect(()=>{sv(SK.l,likes)},[likes]);
  useEffect(()=>{sv(SK.cert,certs)},[certs]);
  useEffect(()=>{sv(SK.arch,archive)},[archive]);

  // Auto-generate certs for closed competitions
  useEffect(()=>{
    const today=new Date().toISOString().slice(0,10);
    const closed=comps.filter(c=>c.closingDate<=today&&!c.archived);
    if(!closed.length)return;
    const newC:Cert[]=[];
    closed.forEach(comp=>{
      const scored=works.filter(w=>w.competitionId===comp.id&&evals.some(e=>e.workId===w.id)).map(w=>{
        const ev=evals.find(e=>e.workId===w.id)!; const stu=users.find(s=>s.id===w.studentId);
        return{work:w,score:tot(ev.scores),stu};
      }).filter(x=>x.stu).sort((a,b)=>b.score-a.score);
      scored.slice(0,20).forEach((it,idx)=>{
        if(!certs.some(x=>x.workId===it.work.id)){
          newC.push({id:mid("cert"),workId:it.work.id,studentId:it.stu!.id,compId:comp.id,compName:comp.nameEn,score:it.score,rank:idx+1,grade:it.stu!.grade,section:it.stu!.section,issuedAt:today});
        }
      });
    });
    if(newC.length)setCerts(p=>[...p,...newC]);
  },[comps,works,evals,users,certs]);

  function handleArchive(comp:Comp, winners:Array<{rank:number;studentName:string;grade:string;section:string;score:number}>){
    const entry:ArchiveEntry={id:mid("arch"),compId:comp.id,compName:comp.nameEn,compNameAr:comp.nameAr,archivedAt:new Date().toISOString().slice(0,10),winners};
    setArchive(p=>[entry,...p]);
    setComps(p=>p.map(c=>c.id===comp.id?{...c,archived:true}:c));
  }

  if(view==="home")return<HomePage onStart={()=>setView("showcase")} onAdmin={()=>setView("adminLogin")} onStudent={()=>setView("studentLogin")}/>;
  if(view==="showcase")return<ShowcasePage comps={comps} users={users} works={works} evals={evals} likes={likes} setLikes={setLikes} setEvals={setEvals} setWorks={setWorks} onBack={()=>setView("home")} onClass={c=>{setSelClass(c);setView("classBoard")}} onParticipate={id=>{setSelCompId(id);setView("studentLogin")}} onWinners={()=>setView("winners")} onAllWorks={()=>setView("allWorks")} />;
  if(view==="allWorks")return<GlobalWorksPage users={users} works={works} evals={evals} comps={comps} likes={likes} setLikes={setLikes} teacher={teacherSess} setEvals={setEvals} setWorks={setWorks} onTeacherLogin={t=>setTeacherSess(t)} onBack={()=>setView("showcase")}/>;
  if(view==="winners")return<WinnersPage comps={comps} works={works} evals={evals} users={users} likes={likes} setLikes={setLikes} adminPw={adminPw} onArchive={handleArchive} onBack={()=>setView("showcase")}/>;
  if(view==="classBoard"&&selClass)return<ClassBoardPage cls={selClass} users={users} works={works} evals={evals} comps={comps} likes={likes} setLikes={setLikes} teacher={teacherSess} setEvals={setEvals} setWorks={setWorks} onTeacherLogin={t=>setTeacherSess(t)} onBack={()=>setView("showcase")}/>;
  if(view==="adminLogin")return<AdminLoginPage adminPw={adminPw} onBack={()=>setView("home")} onOk={()=>setView("admin")}/>;
  if(view==="admin")return<AdminPage users={users} setUsers={setUsers} comps={comps} setComps={setComps} archive={archive} adminPw={adminPw} setAdminPw={setAdminPw} onBack={()=>setView("home")} onTeacher={()=>setView("teacherLogin")}/>;
  if(view==="teacherLogin")return<TeacherLoginPage users={users} onBack={()=>setView("admin")} onOk={t=>{setTeacherSess(t);setView("teacher")}}/>;
  if(view==="teacher"&&teacherSess)return<TeacherPage teacher={teacherSess} users={users} works={works} comps={comps} evals={evals} setEvals={setEvals} onBack={()=>{setTeacherSess(null);setView("admin")}}/>;
  if(view==="studentLogin")return<StudentLoginPage users={users} onBack={()=>setView("showcase")} onOk={s=>{setStudentSess(s);setView("student")}}/>;
  if(view==="student"&&studentSess)return<StudentPage student={studentSess} comps={comps} works={works} evals={evals} certs={certs} selCompId={selCompId} setWorks={setWorks} onBack={()=>{setStudentSess(null);setView("showcase")}}/>;
  return null;
}

/* ════════════════════════════════════
   1. HOME
   ════════════════════════════════════ */
function HomePage({onStart,onAdmin,onStudent}:{onStart:()=>void;onAdmin:()=>void;onStudent:()=>void}){return(
  <div className="relative min-h-screen overflow-hidden bg-[linear-gradient(180deg,#0b1028,#0f1535,#0d1230)] text-white"><Bg/>
    <header className="relative z-20 border-b border-white/10 bg-[#0b1028]/80 backdrop-blur-xl">
      <div className="mx-auto flex max-w-7xl items-center justify-between px-6 py-4">
        <div className="flex items-center gap-4"><Logo/><div><h1 className="text-base font-black uppercase italic tracking-tight text-white sm:text-lg">GWA Arabic & Islamic Champions</h1><p className="text-xs font-semibold text-[#FDB51F]/80" dir="rtl">أبطال اللغة العربية والتربية الإسلامية والدراسات الاجتماعية</p></div></div>
        <div className="flex items-center gap-3">
          <button onClick={onStudent} className="rounded-full border border-white/15 bg-white/5 px-5 py-2.5 text-sm font-bold text-white backdrop-blur hover:bg-white/10 transition">Student Login • دخول الطالب</button>
          <button onClick={onAdmin} className="rounded-full bg-[#FDB51F] px-5 py-2.5 text-sm font-black text-slate-950 shadow-lg shadow-[#FDB51F]/20 hover:bg-[#F7C84A] transition">الإدارة • Admin</button>
        </div>
      </div>
    </header>
    <main className="relative z-10 mx-auto max-w-7xl px-6 py-10 lg:py-14">
      <div className="grid gap-6 lg:grid-cols-2 lg:gap-12 items-center">
        <div className="flex flex-col justify-center order-2 lg:order-1">
          <div className="inline-flex w-fit items-center gap-2 rounded-full border border-[#FDB51F]/25 bg-[#FDB51F]/10 px-5 py-2 mb-5"><span className="h-2 w-2 rounded-full bg-[#FDB51F] animate-pulse"/><span className="text-xs font-black uppercase tracking-[0.22em] text-[#FFE29A]">GEMS World Academy Dubai</span></div>
          <h2 className="text-3xl font-black leading-tight tracking-tight text-white sm:text-4xl xl:text-5xl">Celebrating Our Students' <span className="block text-transparent bg-clip-text bg-gradient-to-r from-[#FDB51F] via-[#1697E8] to-[#10B09B]">Outstanding Achievements</span></h2>
          <p className="mt-4 text-base leading-7 text-white/55 max-w-xl">A world-class bilingual platform showcasing student excellence in Arabic Language, Islamic Education & Social Studies.</p>
          <div className="mt-4 rounded-2xl border border-white/10 bg-white/5 px-5 py-4 backdrop-blur-sm" dir="rtl"><p className="text-base font-bold leading-9 text-white/80 text-right">أهلاً بكم في <span className="text-[#FDB51F] font-black">أكاديمية جيمس العالمية</span> — دبي، حيث نحتفل بإنجازات طلابنا الأبطال المتميزين في:</p>
            <div className="flex flex-wrap gap-2 justify-end mt-3">
              <span className="rounded-full bg-[#1697E8]/20 border border-[#1697E8]/30 px-3 py-1 text-sm font-bold text-sky-200">📖 اللغة العربية</span>
              <span className="rounded-full bg-[#10B09B]/20 border border-[#10B09B]/30 px-3 py-1 text-sm font-bold text-teal-200">🕌 التربية الإسلامية</span>
              <span className="rounded-full bg-[#E21455]/20 border border-[#E21455]/30 px-3 py-1 text-sm font-bold text-pink-200">🌍 الدراسات الاجتماعية</span>
            </div>
          </div>
          <div className="mt-7">
            <button onClick={onStart} className="group rounded-full p-[2px] bg-gradient-to-r from-[#FDB51F] to-[#E21455] shadow-[0_0_40px_rgba(253,181,31,0.35)] hover:shadow-[0_0_60px_rgba(253,181,31,0.5)] transition-all hover:scale-105">
              <div className="rounded-full bg-[#0b1028] px-8 py-4 flex items-center gap-3"><span className="text-xl">🚀</span><span className="text-lg font-black text-white">Start Viewing • بدء المشاهدة</span></div>
            </button>
          </div>
          <div className="mt-8 grid grid-cols-3 gap-3">
            {[{en:"Competitions",num:"✦",ar:"المسابقات",c:"text-[#FDB51F]"},{en:"Class Galleries",num:"✦",ar:"لوحات الصفوف",c:"text-sky-300"},{en:"Criteria",num:"10",ar:"معايير التقييم",c:"text-emerald-300"}].map(s=>(
              <div key={s.en} className="rounded-2xl border border-white/10 bg-white/5 p-4 backdrop-blur-sm text-center"><div className={cn("text-2xl font-black",s.c)}>{s.num}</div><div className="text-[10px] font-bold uppercase tracking-widest text-white/40 mt-1">{s.en}</div><div className="text-xs font-bold text-white/50 mt-1" dir="rtl">{s.ar}</div></div>
            ))}
          </div>
        </div>
        <div className="relative order-1 lg:order-2 flex justify-center"><div className="absolute inset-0 rounded-[3rem] bg-gradient-to-br from-[#1697E8]/20 via-[#FDB51F]/10 to-transparent blur-3xl"/><div className="relative rounded-[2.5rem] overflow-hidden border border-[#FDB51F]/20 shadow-[0_0_80px_rgba(22,151,232,0.2)]"><img src="/hero-student.png" alt="" className="w-full max-w-lg xl:max-w-xl object-cover min-h-[380px]"/><div className="absolute bottom-0 inset-x-0 bg-gradient-to-t from-[#0b1028]/90 via-[#0b1028]/40 to-transparent p-6 flex items-center gap-3"><span className="text-3xl">🏆</span><div><p className="text-lg font-black text-[#FDB51F]">Arabic & Islamic Champions</p><p className="text-sm text-white/60" dir="rtl">أبطال العربية والإسلامية</p></div></div></div></div>
      </div>
    </main>
  </div>
);}

/* ════════════════════════════════════
   2. SHOWCASE
   ════════════════════════════════════ */
function ShowcasePage({comps,users,works,evals,likes,setLikes,setEvals,setWorks,onBack,onParticipate,onWinners}:{comps:Comp[];users:User[];works:Work[];evals:Ev[];likes:Like[];setLikes:React.Dispatch<React.SetStateAction<Like[]>>;setEvals:React.Dispatch<React.SetStateAction<Ev[]>>;setWorks:React.Dispatch<React.SetStateAction<Work[]>>;onBack:()=>void;onClass:(c:ClassInfo)=>void;onParticipate:(id:string)=>void;onWinners:()=>void;onAllWorks:()=>void}){
  const evaluatedCount=useMemo(()=>works.filter(w=>evals.some(e=>e.workId===w.id)).length,[works,evals]);
  const active=useMemo(()=>comps.filter(c=>c.started&&!c.archived),[comps]);
  const[selCompId,setSelCompId]=useState<string|null>(null);
  const[gradeFilter,setGradeFilter]=useState<string|null>(null);
  const selComp=selCompId?comps.find(c=>c.id===selCompId):null;

  // Works for selected competition
  const compWorks=useMemo(()=>{
    if(!selComp)return[];
    let w=works.filter(x=>x.competitionId===selComp.id);
    if(gradeFilter){
      w=w.filter(x=>{const s=users.find(u=>u.id===x.studentId);return s?.grade===gradeFilter});
    }
    return w.sort((a,b)=>b.at.localeCompare(a.at));
  },[selComp,works,users,gradeFilter]);

  // Unique grades that participated in this competition
  const compGrades=useMemo(()=>{
    if(!selComp)return[];
    const grades=new Set<string>();
    works.filter(w=>w.competitionId===selComp.id).forEach(w=>{
      const s=users.find(u=>u.id===w.studentId);
      if(s)grades.add(s.grade);
    });
    return Array.from(grades).sort();
  },[selComp,works,users]);

  // ── COMPETITION ICONS VIEW (no competition selected) ──
  if(!selComp){
    return(
      <div className="relative min-h-screen bg-[linear-gradient(180deg,#0d1230,#111a45,#0d1230)] text-white"><Bg/>
        <header className="sticky top-0 z-20 border-b border-white/10 bg-[#0f1535]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center">
          <div className="flex items-center gap-4"><Logo s/><div><h2 className="text-lg font-black uppercase italic">Active Competitions</h2><p className="text-sm text-white/40" dir="rtl">المسابقات النشطة</p></div></div>
          <div className="flex gap-3">
            <button onClick={onWinners} className="rounded-xl px-5 py-2.5 text-sm font-black bg-gradient-to-r from-[#FDB51F] to-[#E21455] text-white shadow-lg hover:brightness-110">🏅 Winners{evaluatedCount>0?` (${evaluatedCount})`:""}</button>
            <button onClick={onBack} className="rounded-full border border-white/10 bg-white/5 px-5 py-2.5 text-sm font-bold text-white hover:bg-white/10">🏠 Home</button>
          </div>
        </header>
        <main className="max-w-7xl mx-auto px-6 py-10">
          <p className="text-center text-white/50 mb-8 text-lg">Select a competition to view all student submissions • اختر مسابقة لمشاهدة جميع المشاركات</p>
          {active.length===0&&<div className="rounded-3xl border border-white/10 bg-white/5 p-16 text-center"><p className="text-4xl mb-4">🏆</p><p className="text-xl text-white/40">No active competitions • لا توجد مسابقات نشطة</p></div>}
          <div className="grid gap-6 sm:grid-cols-2 lg:grid-cols-3">
            {active.map(c=>{
              const compWorksCount=works.filter(w=>w.competitionId===c.id).length;
              const co=getClassColor(c.nameEn);
              return(
                <button key={c.id} onClick={()=>{setSelCompId(c.id);setGradeFilter(null)}} className="group text-left rounded-[2rem] border border-white/10 p-6 hover:-translate-y-1 transition-all relative overflow-hidden backdrop-blur-sm" style={{background:`linear-gradient(140deg,${co.bg},rgba(15,23,42,0.7))`}}>
                  <div className="absolute top-0 right-0 w-32 h-32 blur-3xl rounded-full transition group-hover:scale-110" style={{background:co.glow}}/>
                  <div className="relative z-10">
                    <div className="flex items-center gap-4 mb-4">
                      <img src={c.logoUrl||"/gems-logo.png"} alt="" className="h-16 w-16 rounded-2xl bg-white p-2 object-contain shadow-lg"/>
                      <div>
                        <h3 className="text-xl font-black text-white leading-tight">{c.nameEn}</h3>
                        <p className="text-sm font-bold text-white/50 mt-0.5" dir="rtl">{c.nameAr}</p>
                      </div>
                    </div>
                    <div className="flex flex-wrap gap-2 mb-4">{c.grades.slice(0,3).map(g=><Pill key={g} t="blue">{g}</Pill>)}{c.grades.length>3&&<Pill t="slate">+{c.grades.length-3}</Pill>}</div>
                    <div className="flex justify-between items-center text-sm font-bold text-white/50 pt-4 border-t border-white/10">
                      <span>🎬 {compWorksCount} Submissions</span>
                      <span className="transition group-hover:text-white" style={{color:co.border}}>View All →</span>
                    </div>
                    <div className="mt-3 flex gap-2">
                      <button onClick={e=>{e.stopPropagation();onParticipate(c.id)}} className="rounded-full bg-[#FDB51F] px-4 py-2 text-xs font-black text-slate-950 hover:bg-[#F7C84A] shadow">Participate 🎓</button>
                      <Pill t="pink">Closes: {c.closingDate}</Pill>
                    </div>
                  </div>
                </button>
              );
            })}
          </div>
        </main>
      </div>
    );
  }

  // ── COMPETITION WORKS VIEW (competition selected) ──
  return(
    <div className="relative min-h-screen bg-[linear-gradient(180deg,#0d1230,#111a45,#0d1230)] text-white"><Bg/>
      <header className="sticky top-0 z-20 border-b border-white/10 bg-[#0f1535]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center">
        <button onClick={()=>{setSelCompId(null);setGradeFilter(null)}} className="rounded-full border border-white/10 bg-white/5 px-5 py-2.5 text-sm font-bold text-white hover:bg-white/10">← All Competitions</button>
        <div className="flex gap-3">
          <button onClick={onWinners} className="rounded-xl px-4 py-2 text-sm font-black bg-gradient-to-r from-[#FDB51F] to-[#E21455] text-white shadow-lg hover:brightness-110">🏅 Winners</button>
          <Logo s/>
        </div>
      </header>

      {/* Competition Header */}
      <section className="relative z-10 border-b border-white/10 bg-[radial-gradient(circle_at_top,rgba(22,151,232,0.2),transparent_35%),linear-gradient(160deg,#111a45,#0d1230)] px-6 py-10">
        <div className="max-w-7xl mx-auto">
          <div className="flex items-center gap-5 mb-6">
            <img src={selComp.logoUrl||"/gems-logo.png"} alt="" className="h-18 w-18 rounded-2xl bg-white p-2 object-contain shadow-lg"/>
            <div>
              <h1 className="text-3xl sm:text-4xl font-black text-white">{selComp.nameEn}</h1>
              <p className="text-xl font-bold text-[#FDB51F]" dir="rtl">{selComp.nameAr}</p>
              <div className="flex gap-2 mt-2"><Pill t="blue">{compWorks.length} Works</Pill><Pill t="pink">Closes: {selComp.closingDate}</Pill></div>
            </div>
          </div>

          {/* Grade Filter Icons */}
          <div className="flex flex-wrap gap-2">
            <button onClick={()=>setGradeFilter(null)} className={cn("rounded-full px-5 py-2.5 text-sm font-black transition border",!gradeFilter?"bg-[#FDB51F] text-slate-950 border-[#FDB51F] shadow-lg":"border-white/10 bg-white/5 text-white/60 hover:bg-white/10")}>
              🌎 All Grades • جميع الصفوف
            </button>
            {compGrades.map(g=>{
              const co=getClassColor(g);
              return(
                <button key={g} onClick={()=>setGradeFilter(g)} className={cn("rounded-full px-5 py-2.5 text-sm font-black transition border",gradeFilter===g?"text-white shadow-lg":"text-white/60 hover:text-white")} style={{borderColor:gradeFilter===g?co.border:"rgba(255,255,255,0.1)",background:gradeFilter===g?co.bg:"rgba(255,255,255,0.03)"}}>
                  {g}
                </button>
              );
            })}
          </div>
        </div>
      </section>

      <main className="max-w-7xl mx-auto px-6 py-10">
        {compWorks.length===0&&<div className="rounded-3xl border border-white/10 bg-white/5 p-16 text-center"><p className="text-xl text-white/40">{gradeFilter?`No works from ${gradeFilter} yet`:"No works submitted yet"}</p></div>}
        <div className="grid gap-8 md:grid-cols-2 xl:grid-cols-3">
          {compWorks.map(w=><WorkCard key={w.id} w={w} users={users} evals={evals} comps={comps} likes={likes} setLikes={setLikes} setEvals={setEvals} setWorks={setWorks}/>)}
        </div>
      </main>
    </div>
  );
}

/* ════════════════════════════════════
   3. CLASS BOARD
   ════════════════════════════════════ */
function ClassBoardPage({cls,users,works,evals,comps,likes,setLikes,teacher,setEvals,setWorks,onTeacherLogin,onBack}:{cls:ClassInfo;users:User[];works:Work[];evals:Ev[];comps:Comp[];likes:Like[];setLikes:React.Dispatch<React.SetStateAction<Like[]>>;teacher:User|null;setEvals:React.Dispatch<React.SetStateAction<Ev[]>>;setWorks:React.Dispatch<React.SetStateAction<Work[]>>;onTeacherLogin:(t:User)=>void;onBack:()=>void}){
  const classWorks=useMemo(()=>works.filter(w=>{const s=users.find(u=>u.id===w.studentId);return s&&s.grade===cls.grade&&s.section===cls.section;}),[cls,users,works]);
  const[showTLogin,setShowTLogin]=useState(false);
  return(
    <div className="relative min-h-screen bg-[#050916] text-white"><Bg/>
      {showTLogin&&<TeacherQuickLogin users={users} onOk={t=>{onTeacherLogin(t);setShowTLogin(false)}} onClose={()=>setShowTLogin(false)}/>}
      <header className="sticky top-0 z-20 border-b border-white/10 bg-[#050916]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center">
        <button onClick={onBack} className="rounded-full border border-white/10 bg-white/5 px-5 py-2.5 text-sm font-bold text-white">← Back</button>
        <div className="flex items-center gap-3">
          {teacher?<Pill t="green">Teacher: {teacher.fullName}</Pill>:<button onClick={()=>setShowTLogin(true)} className="rounded-full bg-sky-500 px-5 py-2 text-sm font-bold text-white hover:bg-sky-400">🔑 Teacher Login</button>}
          <Logo s/>
        </div>
      </header>
      <section className="relative z-10 border-b border-white/10 bg-[radial-gradient(circle_at_top,rgba(22,151,232,0.24),transparent_35%),linear-gradient(160deg,#111a45,#050916_80%)] px-6 py-16 text-center">
        <div className="text-xs font-black uppercase tracking-[0.34em] text-white/40">Gallery of Excellence • معرض التميز</div>
        <h1 className="mt-4 text-5xl sm:text-7xl font-black uppercase tracking-tight">{cls.grade}</h1>
        <div className="mt-2 text-2xl font-black uppercase tracking-[0.28em] text-[#FDB51F]">Section {cls.section}</div>
        <div className="mt-6 flex justify-center gap-3"><Pill t="gold">{cls.count} Students</Pill><Pill t="blue">{classWorks.length} Works</Pill></div>
      </section>
      <main className="max-w-7xl mx-auto px-6 py-12">
        {classWorks.length===0&&<div className="rounded-3xl border border-white/10 bg-white/5 p-16 text-center text-white/40">No works yet • لا توجد أعمال</div>}
        <div className="grid gap-8 md:grid-cols-2 xl:grid-cols-3">
          {classWorks.map(w=><WorkCard key={w.id} w={w} users={users} evals={evals} comps={comps} likes={likes} setLikes={setLikes} setEvals={setEvals} setWorks={setWorks}/>)}
        </div>
      </main>
    </div>
  );
}

/* ════════════════════════════════════
   4. GLOBAL WORKS
   ════════════════════════════════════ */
function GlobalWorksPage({users,works,evals,comps,likes,setLikes,teacher,setEvals,setWorks,onTeacherLogin,onBack}:{users:User[];works:Work[];evals:Ev[];comps:Comp[];likes:Like[];setLikes:React.Dispatch<React.SetStateAction<Like[]>>;teacher:User|null;setEvals:React.Dispatch<React.SetStateAction<Ev[]>>;setWorks:React.Dispatch<React.SetStateAction<Work[]>>;onTeacherLogin:(t:User)=>void;onBack:()=>void}){
  const[showTLogin,setShowTLogin]=useState(false);
  const sorted=useMemo(()=>[...works].sort((a,b)=>b.at.localeCompare(a.at)),[works]);
  return(
    <div className="relative min-h-screen bg-[#0d1230] text-white"><Bg/>
      {showTLogin&&<TeacherQuickLogin users={users} onOk={t=>{onTeacherLogin(t);setShowTLogin(false)}} onClose={()=>setShowTLogin(false)}/>}
      <header className="sticky top-0 z-20 border-b border-white/10 bg-[#0f1535]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center">
        <button onClick={onBack} className="rounded-full border border-white/10 bg-white/5 px-5 py-2.5 text-sm font-bold text-white">← Back</button>
        <div className="flex items-center gap-3">
          {teacher?<Pill t="green">Teacher: {teacher.fullName}</Pill>:<button onClick={()=>setShowTLogin(true)} className="rounded-full bg-sky-500 px-5 py-2 text-sm font-bold text-white hover:bg-sky-400">🔑 Teacher Login</button>}
          <Logo s/>
        </div>
      </header>
      <section className="relative z-10 border-b border-white/10 bg-[radial-gradient(circle_at_top,rgba(253,181,31,0.1),transparent_35%),linear-gradient(160deg,#111a45,#0d1230)] px-6 py-14 text-center">
        <div className="text-xs font-black uppercase tracking-[0.34em] text-[#FDB51F] mb-3">Global Activity Feed • شاشة الجميع</div>
        <h1 className="text-4xl sm:text-6xl font-black uppercase">All Submissions</h1>
        <p className="mt-4 text-lg text-white/50" dir="rtl">جميع مشاركات الطلاب مرتبة حسب تاريخ التحميل</p>
        <div className="mt-6 flex justify-center gap-3"><Pill t="blue">{sorted.length} Works</Pill></div>
      </section>
      <main className="max-w-7xl mx-auto px-6 py-12">
        <div className="grid gap-8 md:grid-cols-2 xl:grid-cols-3">
          {sorted.map(w=><WorkCard key={w.id} w={w} users={users} evals={evals} comps={comps} likes={likes} setLikes={setLikes} setEvals={setEvals} setWorks={setWorks}/>)}
        </div>
      </main>
    </div>
  );
}

/* ════════════════════════════════════
   5. WINNERS PAGE — Competition icons for closed comps only
   ════════════════════════════════════ */
function WinnersPage({comps,works,evals,users,likes,setLikes,adminPw,onArchive,onBack}:{comps:Comp[];works:Work[];evals:Ev[];users:User[];likes:Like[];setLikes:React.Dispatch<React.SetStateAction<Like[]>>;adminPw:string;onArchive:(comp:Comp,winners:any[])=>void;onBack:()=>void}){
  const today=new Date().toISOString().slice(0,10);
  const closedComps=useMemo(()=>comps.filter(c=>c.closingDate<=today&&!c.archived),[comps,today]);
  const[selCompId,setSelCompId]=useState<string|null>(closedComps[0]?.id||null);
  const[archivePw,setArchivePw]=useState("");
  const[archiveErr,setArchiveErr]=useState("");
  const[archiveMode,setArchiveMode]=useState(false);
  const[archiveDone,setArchiveDone]=useState<string|null>(null);

  const selComp=comps.find(c=>c.id===selCompId)||null;

  const rankingsForComp=useMemo(()=>{
    if(!selComp)return[];
    const scored=works.filter(w=>w.competitionId===selComp.id&&evals.some(e=>e.workId===w.id)).map(w=>{
      const ev=evals.find(e=>e.workId===w.id)!; const stu=users.find(s=>s.id===w.studentId);
      return{work:w,score:tot(ev.scores),stu};
    }).filter(x=>x.stu).sort((a,b)=>b.score-a.score).slice(0,20);
    return scored.map((item,i)=>({...item,rank:i+1}));
  },[selComp,works,evals,users]);

  function doArchive(){
    if(archivePw.trim()!==adminPw){setArchiveErr("Incorrect admin password");return}
    if(!selComp)return;
    const winners=rankingsForComp.map(it=>({rank:it.rank,studentName:it.stu!.fullName,grade:it.stu!.grade,section:it.stu!.section,score:it.score}));
    onArchive(selComp,winners);
    setArchiveMode(false); setArchivePw(""); setArchiveErr("");
    setArchiveDone(selComp.nameEn);
    setSelCompId(null);
  }

  return(
    <div className="relative min-h-screen bg-[#0d1230] text-white"><Bg/>
      <div className="pointer-events-none fixed top-0 left-1/2 -translate-x-1/2 w-[600px] h-[300px] bg-[#FDB51F]/8 blur-[100px] rounded-full"/>
      <header className="sticky top-0 z-20 border-b border-white/10 bg-[#0f1535]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center"><button onClick={onBack} className="rounded-full border border-white/10 bg-white/5 px-5 py-2.5 text-sm font-bold text-white">← Back</button><Logo s/></header>

      {/* Hero */}
      <section className="relative z-10 px-6 py-14 text-center border-b border-white/10 bg-[radial-gradient(ellipse_at_top,rgba(253,181,31,0.15),transparent_45%)]">
        <div className="text-5xl mb-4">🏅</div>
        <h1 className="text-4xl sm:text-6xl font-black tracking-tight">Hall of <span className="text-transparent bg-clip-text bg-gradient-to-r from-[#FDB51F] via-yellow-300 to-[#FDB51F]">Champions</span></h1>
        <p className="mt-3 text-xl font-bold text-white/50" dir="rtl">قاعة الفائزين — أيقونات المسابقات المغلقة</p>
        <p className="text-sm text-white/30 mt-2">Only closed competitions appear here • تظهر المسابقات المغلقة فقط</p>
      </section>

      <main className="max-w-6xl mx-auto px-6 py-12">
        {closedComps.length===0&&(
          <div className="rounded-3xl border border-white/10 bg-white/5 p-16 text-center">
            <div className="text-5xl mb-4">⏳</div>
            <p className="text-xl text-white/50">No closed competitions yet</p>
            <p className="text-sm text-white/30 mt-2" dir="rtl">ستظهر المسابقات هنا بعد إغلاقها</p>
          </div>
        )}

        {closedComps.length>0&&(
          <>
            {/* Competition Icons */}
            <div className="mb-10">
              <h2 className="text-xl font-black text-white/70 mb-5 uppercase tracking-widest">Select a Closed Competition • اختر مسابقة مغلقة</h2>
              <div className="flex flex-wrap gap-4">
                {closedComps.map(c=>(
                  <button key={c.id} onClick={()=>{setSelCompId(c.id);setArchiveMode(false);setArchiveDone(null)}}
                    className={cn("group relative rounded-3xl border p-5 flex flex-col items-center gap-3 w-44 transition-all hover:-translate-y-1",selCompId===c.id?"border-[#FDB51F]/50 bg-[#FDB51F]/15 shadow-lg shadow-[#FDB51F]/10":"border-white/10 bg-white/5 hover:border-[#FDB51F]/30")}>
                    <div className="absolute -top-2 -right-2 bg-rose-500 text-white text-[9px] font-black px-2 py-0.5 rounded-full uppercase">Closed</div>
                    <img src={c.logoUrl||"/gems-logo.png"} alt="" className="w-16 h-16 rounded-2xl bg-white object-contain p-1.5 shadow"/>
                    <div className="text-center">
                      <p className="text-sm font-black text-white leading-tight">{c.nameEn}</p>
                      <p className="text-[10px] text-white/40 mt-0.5" dir="rtl">{c.nameAr}</p>
                      <p className="text-[10px] font-bold text-rose-300 mt-1">{c.closingDate}</p>
                    </div>
                    {selCompId===c.id&&<div className="w-2 h-2 rounded-full bg-[#FDB51F]"/>}
                  </button>
                ))}
              </div>
            </div>

            {/* Rankings for selected comp */}
            {selComp&&(
              <div>
                {/* Comp Header */}
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-8 p-6 rounded-[2rem] border border-[#FDB51F]/20 bg-[#FDB51F]/8 backdrop-blur-sm">
                  <div className="flex items-center gap-4">
                    <img src={selComp.logoUrl||"/gems-logo.png"} alt="" className="h-16 w-16 rounded-2xl bg-white object-contain p-1.5"/>
                    <div><h3 className="text-2xl font-black text-white">{selComp.nameEn}</h3><p className="text-base font-bold text-[#FDB51F]" dir="rtl">{selComp.nameAr}</p><div className="flex gap-2 mt-2"><Pill t="pink">Closed: {selComp.closingDate}</Pill><Pill t="gold">{rankingsForComp.length} Ranked</Pill></div></div>
                  </div>
                  <button onClick={()=>setArchiveMode(true)} className="rounded-2xl bg-white/10 border border-white/20 px-5 py-3 text-sm font-black text-white hover:bg-white/15 transition flex items-center gap-2 shrink-0">
                    🗄️ Send to Archive • إرسال للأرشيف
                  </button>
                </div>

                {/* Archive Password Modal */}
                {archiveMode&&(
                  <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 p-4 backdrop-blur-sm">
                    <div className="w-full max-w-sm rounded-3xl border border-white/10 bg-[#0f1535] p-6 shadow-2xl">
                      <h3 className="text-xl font-black text-white mb-2">🗄️ Archive Competition</h3>
                      <p className="text-sm text-white/50 mb-5">Enter admin password to archive <b className="text-[#FDB51F]">{selComp.nameEn}</b></p>
                      <div className="space-y-4">
                        <Inp type="password" value={archivePw} onChange={e=>{setArchivePw(e.target.value);setArchiveErr("")}} placeholder="Admin password"/>
                        {archiveErr&&<div className="rounded-xl bg-rose-500/10 border border-rose-400/20 px-4 py-3 text-sm text-rose-200 font-bold">{archiveErr}</div>}
                        <div className="flex gap-3">
                          <button onClick={()=>{setArchiveMode(false);setArchivePw("");setArchiveErr("")}} className="flex-1 rounded-2xl border border-white/10 bg-white/5 py-3 text-sm font-bold text-white">Cancel</button>
                          <button onClick={doArchive} className="flex-1 rounded-2xl bg-[#FDB51F] py-3 text-sm font-black text-slate-950">Archive ✓</button>
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {archiveDone&&<div className="mb-6 rounded-2xl border border-emerald-400/20 bg-emerald-500/10 px-5 py-3 text-emerald-200 font-bold">✅ "{archiveDone}" has been archived successfully!</div>}

                {rankingsForComp.length===0?(
                  <div className="rounded-3xl border border-white/10 bg-white/5 p-12 text-center text-white/40">No evaluated works yet for this competition</div>
                ):(
                  <div className="space-y-6">
                    {rankingsForComp.map(({work,score,stu,rank})=>{
                      const wl=likes.filter(l=>l.workId===work.id);
                      const isTop3=rank<=3;
                      return(
                        <article key={work.id} className={cn("rounded-[2rem] border overflow-hidden",rank===1?"border-[#FDB51F]/40 bg-gradient-to-br from-[#FDB51F]/10 via-[#1F2E76]/50 to-[#0d1230] shadow-lg shadow-[#FDB51F]/10":rank===2?"border-slate-400/25 bg-[#1F2E76]/30":rank===3?"border-amber-700/25 bg-[#1F2E76]/25":"border-white/10 bg-white/5")}>
                          <div className="grid lg:grid-cols-[380px_1fr]">
                            <div className="relative"><div className="absolute top-4 left-4 z-10 flex items-center gap-2"><div className={cn("rounded-2xl px-4 py-2 text-2xl font-black shadow-lg backdrop-blur-md",rank===1?"bg-[#FDB51F] text-slate-900":rank===2?"bg-slate-300 text-slate-900":rank===3?"bg-amber-700 text-white":"bg-white/10 text-white border border-white/15")}>{MEDALS[rank-1]||`#${rank}`}</div>{isTop3&&<div className="rounded-full bg-white/10 backdrop-blur px-3 py-1 text-xs font-black text-white/80">#{rank} Champion</div>}</div><video src={work.videoUrl} poster={work.thumb} controls className="w-full h-full object-cover min-h-[200px]"/></div>
                            <div className="p-6 flex flex-col justify-between">
                              <div><h3 className="text-xl font-black text-white">{work.title}</h3><div className="mt-3 flex items-center gap-3"><div className={cn("flex h-12 w-12 items-center justify-center rounded-2xl text-lg font-black shrink-0",rank===1?"bg-[#FDB51F] text-slate-900":rank===2?"bg-slate-300 text-slate-900":rank===3?"bg-amber-700 text-white":"bg-[#1F2E76] text-[#FDB51F]")}>{stu!.fullName[0]}</div><div><p className="text-base font-black text-white">{stu!.fullName}</p><p className="text-sm text-white/40">{stu!.grade} • Section {stu!.section}</p></div></div></div>
                              <div className="mt-4 space-y-3"><div className="flex items-center gap-3"><div className={cn("rounded-2xl px-5 py-3 flex items-center gap-3",rank===1?"bg-[#FDB51F]/15 border border-[#FDB51F]/30":"bg-emerald-500/10 border border-emerald-400/20")}><span className="text-xs font-bold text-white/50">Score</span><span className={cn("text-3xl font-black",rank===1?"text-[#FDB51F]":"text-emerald-300")}>{score}</span><span className="text-white/30">/100</span></div></div>
                              <div className="flex flex-wrap gap-3 text-xs font-bold"><span className="text-sky-300">👫 {wl.filter(l=>l.type==="friend").length}</span><span className="text-emerald-300">👨‍👩‍👧 {wl.filter(l=>l.type==="parent").length}</span><span className="text-pink-300">👩‍🏫 {wl.filter(l=>l.type==="teacher").length}</span><span className="text-white/30">❤️ {wl.length} Total</span></div>
                              <LikesPanel workId={work.id} likes={likes} setLikes={setLikes}/></div>
                            </div>
                          </div>
                        </article>
                      );
                    })}
                  </div>
                )}
              </div>
            )}
          </>
        )}
      </main>
    </div>
  );
}

/* ════════════════════════════════════
   6. AUTH PAGES
   ════════════════════════════════════ */
function AdminLoginPage({adminPw,onBack,onOk}:{adminPw:string;onBack:()=>void;onOk:()=>void}){
  const[pw,setPw]=useState(""); const[err,setErr]=useState(""); const[showPw,setShowPw]=useState(false);
  function sub(e:React.FormEvent){e.preventDefault(); const t=pw.trim(); if(t===adminPw){setErr("");onOk()}else{setErr("Incorrect password");setPw("")}}
  return(
    <div className="relative min-h-screen bg-[#0d1230] flex items-center justify-center px-6"><Bg/>
      <div className="relative w-full max-w-md rounded-[2.4rem] border border-white/10 bg-white/5 p-8 backdrop-blur-xl shadow-2xl"><button onClick={onBack} className="text-white/50 text-sm font-bold mb-6 hover:text-white">← Back</button><div className="flex flex-col items-center mb-8"><Logo/><h2 className="mt-5 text-2xl font-black text-white">Administration Portal</h2><p className="text-sm text-white/50 mt-1">بوابة الإدارة</p></div>
        <form onSubmit={sub} className="space-y-5"><div><FL>Admin Password</FL><div className="relative"><Inp type={showPw?"text":"password"} value={pw} onChange={e=>{setPw(e.target.value);if(err)setErr("")}} className="pr-14" autoComplete="off"/><button type="button" onClick={()=>setShowPw(s=>!s)} className="absolute right-4 top-1/2 -translate-y-1/2 text-white/40 text-xs font-bold">{showPw?"Hide":"Show"}</button></div></div>
          {err&&<div className="rounded-2xl border border-rose-400/20 bg-rose-500/10 px-4 py-3 text-sm font-bold text-rose-200 text-center">⚠️ {err}</div>}
          <button type="submit" className="w-full rounded-2xl bg-[#FDB51F] px-5 py-4 text-base font-black text-slate-950 shadow-lg shadow-[#FDB51F]/20 hover:bg-[#F7C84A]">دخول الإدارة • Enter</button></form></div>
    </div>
  );
}

function TeacherLoginPage({users,onBack,onOk}:{users:User[];onBack:()=>void;onOk:(t:User)=>void}){
  const[u,setU]=useState(""); const[p,setP]=useState(""); const[err,setErr]=useState("");
  function sub(e:React.FormEvent){e.preventDefault(); const t=users.find(x=>x.role==="teacher"&&x.username===u.trim()&&x.password===p&&x.active); if(t)onOk(t); else setErr("Invalid credentials or deactivated account")}
  return(
    <div className="relative min-h-screen bg-[#0d1230] flex items-center justify-center px-6"><Bg/>
      <div className="relative w-full max-w-md rounded-[2.4rem] border border-white/10 bg-white/5 p-8 backdrop-blur-xl shadow-2xl"><button onClick={onBack} className="text-white/50 text-sm font-bold mb-6 hover:text-white">← Back</button><div className="flex flex-col items-center mb-8"><Logo/><h2 className="mt-5 text-2xl font-black text-white">Teacher Login</h2></div>
        <form onSubmit={sub} className="space-y-5"><div><FL>Username</FL><Inp value={u} onChange={e=>setU(e.target.value)} placeholder="heba.6x"/></div><div><FL>Password</FL><Inp type="password" value={p} onChange={e=>setP(e.target.value)}/></div>
          {err&&<div className="rounded-2xl border border-rose-400/20 bg-rose-500/10 px-4 py-3 text-sm text-rose-200 font-bold">{err}</div>}
          <button type="submit" className="w-full rounded-2xl bg-sky-500 px-5 py-4 text-base font-black text-white shadow-lg hover:bg-sky-400">Sign In</button></form></div>
    </div>
  );
}

function StudentLoginPage({users,onBack,onOk}:{users:User[];onBack:()=>void;onOk:(s:User)=>void}){
  const[u,setU]=useState(""); const[p,setP]=useState(""); const[err,setErr]=useState("");
  function sub(e:React.FormEvent){e.preventDefault(); const s=users.find(x=>x.role==="student"&&x.username===u.trim()&&x.password===p&&x.active); if(s)onOk(s); else setErr("Invalid credentials or deactivated account")}
  return(
    <div className="relative min-h-screen bg-[#0d1230] flex items-center justify-center px-6"><Bg/>
      <div className="relative w-full max-w-md rounded-[2.4rem] border border-white/10 bg-white/5 p-8 backdrop-blur-xl shadow-2xl"><button onClick={onBack} className="text-white/50 text-sm font-bold mb-6 hover:text-white">← Back</button><div className="flex flex-col items-center mb-8"><Logo/><h2 className="mt-5 text-2xl font-black text-white">Student Login • دخول الطالب</h2></div>
        <form onSubmit={sub} className="space-y-5"><div><FL>Username • اسم المستخدم</FL><Inp value={u} onChange={e=>setU(e.target.value)} placeholder="mariam.salem"/></div><div><FL>Password • كلمة المرور</FL><Inp type="password" value={p} onChange={e=>setP(e.target.value)}/></div>
          {err&&<div className="rounded-2xl border border-rose-400/20 bg-rose-500/10 px-4 py-3 text-sm text-rose-200 font-bold">{err}</div>}
          <button type="submit" className="w-full rounded-2xl bg-[#FDB51F] px-5 py-4 text-base font-black text-slate-950 shadow-lg shadow-[#FDB51F]/20 hover:bg-[#F7C84A]">Login 🎓</button></form></div>
    </div>
  );
}

/* ════════════════════════════════════
   7. ADMIN PAGE (with Archive tab)
   ════════════════════════════════════ */
function AdminPage({users,setUsers,comps,setComps,archive,adminPw,setAdminPw,onBack,onTeacher}:{users:User[];setUsers:React.Dispatch<React.SetStateAction<User[]>>;comps:Comp[];setComps:React.Dispatch<React.SetStateAction<Comp[]>>;archive:ArchiveEntry[];adminPw:string;setAdminPw:(pw:string)=>void;onBack:()=>void;onTeacher:()=>void}){
  const[showPwModal,setShowPwModal]=useState(false);
  const[newPw,setNewPw]=useState("");
  const teachers=useMemo(()=>users.filter(u=>u.role==="teacher"),[users]);
  const[tab,setTab]=useState<"users"|"export"|"comps"|"archive">("users");
  const fRef=useRef<HTMLInputElement>(null);
  const[name,setName]=useState(""); const[role,setRole]=useState<Role>("student"); const[grade,setGrade]=useState("Grade 6"); const[sec,setSec]=useState("6X"); const[tid,setTid]=useState(""); const[umsg,setUmsg]=useState("");
  const[nAr,setNAr]=useState(""); const[nEn,setNEn]=useState(""); const[cGrades,setCGrades]=useState<string[]>(["Grade 6"]); const[rAr,setRAr]=useState(""); const[rEn,setREn]=useState(""); const[cClose,setCClose]=useState("2026-06-01"); const[cCert,setCCert]=useState("2026-06-15"); const[cmsg,setCmsg]=useState(""); const[impMsg,setImpMsg]=useState("");
  const[compCriteria,setCompCriteria]=useState<string[]>([...DEFAULT_CRITERIA]);
  const[selPreset,setSelPreset]=useState<string>("");

  function applyPreset(idx:string){
    setSelPreset(idx);
    if(idx==="custom"){setNAr("");setNEn("");setCompCriteria([...DEFAULT_CRITERIA]);return}
    const p=COMP_PRESETS[Number(idx)];
    if(p){setNAr(p.nameAr);setNEn(p.nameEn);setCompCriteria([...p.criteria])}
  }

  function createUser(e:React.FormEvent){e.preventDefault();if(!name.trim()){setUmsg("Enter name");return}if(role==="student"&&!tid){setUmsg("Select teacher");return}
    const u=mkUser(name,role,grade,sec,role==="student"?tid:undefined,users);setUsers(p=>[u,...p]);setUmsg(`✅ ${u.username} / ${u.password}`);setName("")}
  function createComp(e:React.FormEvent){e.preventDefault();if(!nAr.trim()||!nEn.trim()){setCmsg("Enter both names");return}
    const code=Math.random().toString(36).slice(2,8).toUpperCase();
    setComps(p=>[{id:mid("c"),nameAr:nAr,nameEn:nEn,logoUrl:"/gems-logo.png",grades:cGrades,rulesAr:rAr,rulesEn:rEn,closingDate:cClose,certDate:cCert,started:true,archived:false,criteria:compCriteria,evalCode:code},...p]);setCmsg(`✅ Created • Evaluation Code: ${code}`);setNAr("");setNEn("");setRAr("");setREn("")}
  function findCol(h:string[],p:string[]){for(let i=0;i<h.length;i++){const s=h[i].toLowerCase().replace(/[\s_\-().]+/g,"");for(const x of p)if(s===x||s.includes(x))return i}return -1}
  async function handleExcelUpload(e:React.ChangeEvent<HTMLInputElement>){const file=e.target.files?.[0];if(!file)return;
    try{const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array"});const raw=XLSX.utils.sheet_to_json<(string|number)[]>(wb.Sheets[wb.SheetNames[0]],{header:1,defval:""});if(raw.length<2){setImpMsg("⚠️ Empty file");return}
      const headers=raw[0].map(h=>String(h));const rows=raw.slice(1);const ni=findCol(headers,["name","fullname","الاسم"]),ri=findCol(headers,["role","الدور"]),gi=findCol(headers,["grade","الصف"]),si=findCol(headers,["section","القسم","sec"]);const ai=ni>=0?ni:0;const imported:User[]=[];
      rows.forEach(r=>{const n=String(r[ai]||"").trim();if(!n||n.toLowerCase()==="name")return;const rl:Role=["teacher","معلم"].includes(String(r[ri]||"").toLowerCase())?"teacher":"student";imported.push(mkUser(n,rl,String(r[gi]||"Grade 6").trim(),String(r[si]||"6X").trim(),undefined,[...users,...imported]))});
      if(!imported.length){setImpMsg("⚠️ No valid rows");return}setUsers(p=>[...imported,...p]);setImpMsg(`✅ Imported ${imported.length} users`);e.target.value=""}catch{setImpMsg("⚠️ Error reading Excel")}}
  function downloadExcel(){const data=users.map(u=>({Name:u.fullName,Role:u.role,Grade:u.grade,Section:u.section,Username:u.username,Password:u.password,Active:u.active}));const ws=XLSX.utils.json_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Users");XLSX.writeFile(wb,"gwa-users.xlsx")}
  function downloadTemplate(){const t=[{Name:"Mariam Salem",Role:"Student",Grade:"Grade 6",Section:"6X"}];const ws=XLSX.utils.json_to_sheet(t);ws["!cols"]=[{wch:25},{wch:10},{wch:12},{wch:10}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Users");XLSX.writeFile(wb,"gwa-template.xlsx")}
  async function generateCards(){
    const qrData:Record<string,string>={};for(const u of users){try{qrData[u.id]=await QRCode.toDataURL(`GWA|${u.fullName}|${u.username}|${u.password}`,{width:140,margin:1})}catch{qrData[u.id]=""}}
    const win=window.open("","_blank");if(!win)return;
    const cards=users.map((u,i)=>{const isT=u.role==="teacher";const qr=qrData[u.id]||"";return`<div class="card"><div class="ci"><div class="ch"><div class="hl"><img src="/gems-logo.png" class="logo"/><div><div class="pen">GWA Arabic & Islamic Champions</div><div class="par" style="color:${isT?"#1697E8":"#FDB51F"}">${isT?"Teacher":"Student"} Login</div></div></div>${qr?`<img src="${qr}" class="qr"/>`:""}</div><hr class="div"/><div class="inf"><div>Name : <b>${u.fullName}</b></div><div>ID : <b>${i+1}</b></div><div>Grade : <b>${u.grade}</b></div><div>Section : <b>${u.section}</b></div></div><div class="badge" style="background:${isT?"#1697E8":"#FDB51F"};color:${isT?"#fff":"#111"}">${u.role}</div><div class="creds"><div class="url">gwa.gems-world-academy.ae</div><div>• Username: <span class="hl2">${u.username}</span></div><div>• Password: <span class="hl2">${u.password}</span></div></div></div></div>`}).join("");
    win.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8"/><title>GWA Cards</title><style>*{box-sizing:border-box;margin:0;padding:0}body{background:#f0f0f0;font-family:Arial,sans-serif;padding:24px}.tb{display:flex;gap:12px;justify-content:center;margin-bottom:24px;padding:16px;background:#1F2E76;border-radius:16px}.tb button{padding:10px 24px;border-radius:999px;border:none;font-weight:700;cursor:pointer;background:#FDB51F}.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;max-width:1100px;margin:0 auto}.card{break-inside:avoid}.ci{background:#fff;border:3px solid #5ECEC8;border-radius:20px;padding:20px;min-height:220px}.ch{display:flex;justify-content:space-between;margin-bottom:12px}.hl{display:flex;align-items:center;gap:10px}.logo{width:60px;object-fit:contain}.pen{font-size:14px;font-weight:800;color:#1F2E76}.par{font-size:11px;font-weight:700;margin-top:2px}.qr{width:90px}.div{border:none;border-top:1.5px solid #eee;margin:10px 0}.inf{font-size:13px;line-height:1.8}.badge{display:inline-block;border-radius:999px;padding:4px 14px;font-size:11px;font-weight:800;margin:10px 0}.creds{background:#fafafa;border-radius:10px;padding:10px;border:1px solid #eee;font-size:12px;line-height:1.9}.url{font-size:10px;color:#777;margin-bottom:4px}.hl2{background:#FDB51F;padding:1px 6px;border-radius:4px;font-weight:700}@media print{body{background:#fff;padding:8px}.tb{display:none}}</style></head><body><div class="tb"><button onclick="window.print()">🖨 Print</button><button onclick="window.close()">Close</button></div><div class="grid">${cards}</div></body></html>`);win.document.close()}

  return(
    <div className="relative min-h-screen bg-[#0d1230] text-white"><Bg/>
      <header className="sticky top-0 z-20 border-b border-white/10 bg-[#0f1535]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center">
        <div className="flex items-center gap-4"><Logo s/><div><div className="text-[10px] font-bold uppercase tracking-[0.28em] text-[#FDB51F]">Administration Portal • بوابة الإدارة</div><h1 className="text-xl font-black">User & Competition Management</h1></div></div>
        <div className="flex gap-3">
          <button onClick={()=>{setNewPw(adminPw);setShowPwModal(true)}} className="rounded-full border border-white/10 bg-white/5 px-5 py-2 text-sm font-bold text-white hover:bg-white/10">🔑 Password</button>
          <button onClick={onTeacher} className="rounded-full bg-sky-500 px-5 py-2 text-sm font-bold text-white hover:bg-sky-400">Teacher Login</button>
          <button onClick={onBack} className="rounded-full border border-white/10 bg-white/5 px-5 py-2 text-sm font-bold text-white hover:bg-white/10">Home</button>
        </div>
      </header>

      {showPwModal&&(
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 p-4 backdrop-blur-sm">
          <div className="w-full max-w-sm rounded-3xl border border-white/10 bg-[#0f1535] p-6 shadow-2xl">
            <h3 className="text-xl font-black text-white mb-2">Change Admin Password</h3>
            <p className="text-sm text-white/50 mb-5">تغيير كلمة سر الإدارة</p>
            <div className="space-y-4">
              <Inp value={newPw} onChange={e=>setNewPw(e.target.value)} placeholder="New Password"/>
              <div className="flex gap-3">
                <button onClick={()=>setShowPwModal(false)} className="flex-1 rounded-2xl border border-white/10 bg-white/5 py-3 text-sm font-bold text-white/60">Cancel</button>
                <button onClick={()=>{setAdminPw(newPw.trim());setShowPwModal(false)}} className="flex-1 rounded-2xl bg-[#FDB51F] py-3 text-sm font-black text-slate-950">Save ✓</button>
              </div>
            </div>
          </div>
        </div>
      )}
      <main className="max-w-6xl mx-auto px-6 py-8">
        <div className="flex flex-wrap gap-2 mb-8 bg-white/5 border border-white/10 rounded-2xl p-1.5 overflow-x-auto">
          {([["users","👥 Users"],["export","📊 Excel & Cards"],["comps","🏆 Competitions"],["archive","🗄️ Archive • الأرشيف"]] as const).map(([k,l])=>(
            <button key={k} onClick={()=>setTab(k)} className={cn("rounded-xl px-5 py-2.5 text-sm font-black whitespace-nowrap transition",tab===k?"bg-[#FDB51F] text-slate-950":"text-white/50 hover:text-white")}>{l}</button>
          ))}
        </div>

        {tab==="users"&&(
          <div className="grid gap-8 lg:grid-cols-2">
            <div className="space-y-5">
              <div className="rounded-[2rem] border border-white/10 bg-white/5 p-6 backdrop-blur-sm">
                <h2 className="text-xl font-black mb-4">Create Account</h2>
                <form onSubmit={createUser} className="space-y-4">
                  <div><FL>Full Name</FL><Inp value={name} onChange={e=>setName(e.target.value)} placeholder="Name"/></div>
                  <div className="grid grid-cols-3 gap-3"><div><FL>Role</FL><Sel value={role} onChange={e=>setRole(e.target.value as Role)}><option value="student">Student</option><option value="teacher">Teacher</option></Sel></div><div><FL>Grade</FL><Sel value={grade} onChange={e=>setGrade(e.target.value)}>{GRADES.map(g=><option key={g} value={g}>{g}</option>)}</Sel></div><div><FL>Section</FL><Sel value={sec} onChange={e=>setSec(e.target.value)}>{SECTIONS.map(s=><option key={s} value={s}>{s}</option>)}</Sel></div></div>
                  {role==="student"&&<div><FL>Assigned Teacher</FL><Sel value={tid} onChange={e=>setTid(e.target.value)}><option value="">Select…</option>{teachers.map(t=><option key={t.id} value={t.id}>{t.fullName} ({t.grade}/{t.section})</option>)}</Sel></div>}
                  {umsg&&<div className="rounded-2xl border border-emerald-400/20 bg-emerald-500/10 px-4 py-3 text-sm text-emerald-200">{umsg}</div>}
                  <button type="submit" className="w-full rounded-2xl bg-sky-500 text-white font-black py-3.5 hover:bg-sky-400 shadow-lg">Save Account</button>
                </form>
              </div>
              <div className="rounded-[2rem] border border-white/10 bg-white/5 p-6 backdrop-blur-sm">
                <h2 className="text-xl font-black mb-1">📤 Upload Excel</h2>
                <div className="rounded-2xl border-2 border-dashed border-white/10 bg-black/15 p-5 text-center mt-4"><input ref={fRef} type="file" accept=".xlsx,.xls" onChange={handleExcelUpload} className="hidden"/><button type="button" onClick={()=>fRef.current?.click()} className="rounded-2xl bg-white/10 px-6 py-3 text-sm font-black text-white hover:bg-white/15 w-full">📁 Choose Excel File</button></div>
                {impMsg&&<div className="mt-3 rounded-2xl border border-emerald-400/20 bg-emerald-500/10 px-4 py-3 text-sm text-emerald-200">{impMsg}</div>}
                <button type="button" onClick={downloadTemplate} className="mt-3 w-full rounded-2xl border border-white/10 bg-white/5 px-4 py-3 text-sm font-bold text-white/60 hover:bg-white/10">📋 Download Template</button>
              </div>
            </div>
            <div className="rounded-[2rem] border border-white/10 bg-white/5 p-6 backdrop-blur-sm max-h-[800px] overflow-y-auto">
              <h2 className="text-xl font-black mb-4">All Accounts ({users.length})</h2>
              <div className="space-y-3">{users.map(u=>(
                <div key={u.id} className={cn("rounded-2xl border p-4 transition",u.active?"border-white/10 bg-black/20":"border-rose-500/20 bg-rose-500/5 opacity-60")}>
                  <div className="flex justify-between items-start gap-2">
                    <div className="min-w-0"><div className="font-black text-white flex items-center gap-2 flex-wrap">{u.fullName}{!u.active&&<span className="text-[9px] bg-rose-500 text-white px-2 py-0.5 rounded-full uppercase">Disabled</span>}</div><div className="flex flex-wrap gap-1.5 mt-2"><Pill t={u.role==="teacher"?"green":"blue"}>{u.role}</Pill><Pill t="gold">{u.grade}</Pill><Pill t="pink">{u.section}</Pill></div><div className="mt-1.5 text-xs text-white/40 font-bold">👤 {u.username} / 🔑 {u.password}</div></div>
                    <button onClick={()=>setUsers(prev=>prev.map(it=>it.id===u.id?{...it,active:!it.active}:it))} className={cn("rounded-xl px-3 py-2 text-xs font-black transition shrink-0",u.active?"bg-rose-500 text-white hover:bg-rose-600":"bg-emerald-500 text-white hover:bg-emerald-600")}>{u.active?"Disable":"Activate"}</button>
                  </div>
                </div>
              ))}</div>
            </div>
          </div>
        )}

        {tab==="export"&&(
          <div className="grid gap-6 lg:grid-cols-2">
            <div className="rounded-[2rem] border border-white/10 bg-white/5 p-6 backdrop-blur-sm"><h2 className="text-xl font-black mb-2">📊 Download Excel</h2><p className="text-sm text-white/50 mb-4">Export all accounts with credentials</p><button onClick={downloadExcel} className="w-full rounded-2xl bg-emerald-500 text-white font-black py-4 hover:bg-emerald-400 shadow-lg">📥 Download Users Excel</button></div>
            <div className="rounded-[2rem] border border-white/10 bg-white/5 p-6 backdrop-blur-sm"><h2 className="text-xl font-black mb-2">🪪 Educational Cards PDF</h2><p className="text-sm text-white/50 mb-4">Printable cards with QR, logo, credentials</p><button onClick={generateCards} className="w-full rounded-2xl bg-[#FDB51F] text-slate-950 font-black py-4 hover:bg-[#F7C84A] shadow-lg">🪪 Generate & Download Cards</button></div>
          </div>
        )}

        {tab==="comps"&&(
          <div className="grid gap-8 lg:grid-cols-2">
            <div className="rounded-[2rem] border border-white/10 bg-white/5 p-6 backdrop-blur-sm">
              <h2 className="text-xl font-black mb-4">Create Competition</h2>
              <form onSubmit={createComp} className="space-y-4">
                {/* Preset Dropdown */}
                <div><FL>Competition Template • قالب المسابقة</FL>
                  <Sel value={selPreset} onChange={e=>applyPreset(e.target.value)}>
                    <option value="">— Choose a preset or create custom —</option>
                    {COMP_PRESETS.map((p,i)=><option key={i} value={String(i)}>{p.nameEn} — {p.nameAr}</option>)}
                    <option value="custom">✏️ Custom Competition — مسابقة مخصصة</option>
                  </Sel>
                </div>
                <div className="grid grid-cols-2 gap-3"><div><FL>Arabic Name</FL><Inp value={nAr} onChange={e=>setNAr(e.target.value)} dir="rtl" placeholder="مسابقة..."/></div><div><FL>English Name</FL><Inp value={nEn} onChange={e=>setNEn(e.target.value)} placeholder="Competition..."/></div></div>
                <div><FL>Grades</FL><div className="flex flex-wrap gap-2">{GRADES.map(g=><button key={g} type="button" onClick={()=>setCGrades(p=>p.includes(g)?p.filter(x=>x!==g):[...p,g])} className={cn("rounded-full px-3 py-1.5 text-xs font-bold border transition",cGrades.includes(g)?"bg-[#FDB51F] text-slate-950 border-[#FDB51F]":"bg-white/5 text-white/50 border-white/10")}>{g}</button>)}</div></div>
                <div><FL>Rules (Arabic)</FL><TA rows={3} value={rAr} onChange={e=>setRAr(e.target.value)} dir="rtl"/></div>
                <div><FL>Rules (English)</FL><TA rows={3} value={rEn} onChange={e=>setREn(e.target.value)}/></div>
                <div className="grid grid-cols-2 gap-3"><div><FL>Closing Date</FL><Inp type="date" value={cClose} onChange={e=>setCClose(e.target.value)}/></div><div><FL>Certs Date</FL><Inp type="date" value={cCert} onChange={e=>setCCert(e.target.value)}/></div></div>
                {cmsg&&<div className="rounded-2xl border border-emerald-400/20 bg-emerald-500/10 px-4 py-3 text-sm text-emerald-200">{cmsg}</div>}
                <button type="submit" className="w-full rounded-2xl bg-[#FDB51F] text-slate-950 font-black py-3.5 hover:bg-[#F7C84A] shadow-lg">Save Competition 🏆</button>
              </form>
            </div>
            <div className="rounded-[2rem] border border-white/10 bg-white/5 p-6 max-h-[700px] overflow-y-auto">
              <h2 className="text-xl font-black mb-4">All Competitions ({comps.length})</h2>
              <div className="space-y-4">{comps.map(c=>(
                <div key={c.id} className={cn("rounded-2xl border p-5",c.archived?"border-white/5 bg-white/2 opacity-40":"border-white/10 bg-black/20")}>
                  <div className="flex items-center gap-3 mb-2"><img src={c.logoUrl||"/gems-logo.png"} alt="" className="h-10 w-10 rounded-lg bg-white p-1 object-contain"/><div><div className="font-black text-white">{c.nameEn} {c.archived&&<span className="text-[9px] bg-white/20 text-white px-2 py-0.5 rounded-full ml-1">Archived</span>}</div><div className="text-sm text-white/40" dir="rtl">{c.nameAr}</div></div></div>
                  <div className="flex flex-wrap gap-1.5">{c.grades.map(g=><Pill key={g} t="blue">{g}</Pill>)}<Pill t="pink">Closes: {c.closingDate}</Pill></div>
                  <div className="mt-3 flex items-center gap-2">
                    <span className="text-xs text-white/40">🔐 Eval Code:</span>
                    <span className="text-sm font-black text-[#FDB51F] tracking-[0.2em]">{c.evalCode||"—"}</span>
                    <button onClick={()=>{const nc=prompt("New evaluation code:",c.evalCode||"");if(nc!==null)setComps(p=>p.map(x=>x.id===c.id?{...x,evalCode:nc.toUpperCase()}:x))}} className="text-xs text-sky-300 hover:text-sky-200 font-bold ml-2">Change</button>
                  </div>
                </div>
              ))}</div>
            </div>
          </div>
        )}

        {tab==="archive"&&(
          <div>
            <h2 className="text-2xl font-black mb-2 text-white">🗄️ Competitions Archive • أرشيف المسابقات</h2>
            <p className="text-sm text-white/50 mb-8">Competitions are automatically added here when winners are sent to archive from the Winners page.</p>
            {archive.length===0?(
              <div className="rounded-3xl border border-white/10 bg-white/5 p-16 text-center"><p className="text-4xl mb-3">🗄️</p><p className="text-white/40">No archived competitions yet</p><p className="text-sm text-white/25 mt-2">Send a competition to archive from the Winners page → Archive button</p></div>
            ):(
              <div className="space-y-6">{archive.map(entry=>(
                <div key={entry.id} className="rounded-[2rem] border border-[#FDB51F]/20 bg-gradient-to-br from-[#FDB51F]/8 to-[#1F2E76]/30 p-6 backdrop-blur-sm">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-5">
                    <div><h3 className="text-xl font-black text-white">{entry.compName}</h3><p className="text-base font-bold text-[#FDB51F]" dir="rtl">{entry.compNameAr}</p><p className="text-xs text-white/40 mt-1">Archived: {entry.archivedAt}</p></div>
                    <Pill t="gold">{entry.winners.length} Winners</Pill>
                  </div>
                  <div className="rounded-2xl border border-white/10 bg-black/20 overflow-hidden">
                    <table className="w-full text-sm"><thead><tr className="border-b border-white/10 text-white/40 text-[10px] uppercase tracking-widest"><th className="text-left p-3">Rank</th><th className="text-left p-3">Student</th><th className="text-left p-3">Grade</th><th className="text-left p-3">Section</th><th className="text-right p-3">Score</th></tr></thead>
                      <tbody>{entry.winners.map((w,i)=><tr key={i} className="border-b border-white/5 hover:bg-white/5 transition"><td className="p-3 font-black text-white">{MEDALS[w.rank-1]||`#${w.rank}`}</td><td className="p-3 font-bold text-white">{w.studentName}</td><td className="p-3 text-white/60">{w.grade}</td><td className="p-3 text-white/60">{w.section}</td><td className="p-3 text-right font-black text-[#FDB51F]">{w.score}/100</td></tr>)}</tbody>
                    </table>
                  </div>
                </div>
              ))}</div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}

/* ════════════════════════════════════
   8. TEACHER PAGE
   ════════════════════════════════════ */
function TeacherPage({teacher,users,works,comps,evals,setEvals,onBack}:{teacher:User;users:User[];works:Work[];comps:Comp[];evals:Ev[];setEvals:React.Dispatch<React.SetStateAction<Ev[]>>;onBack:()=>void}){
  const myStudents=useMemo(()=>users.filter(u=>u.teacherId===teacher.id),[teacher.id,users]);
  const myWorks=useMemo(()=>works.filter(w=>myStudents.some(s=>s.id===w.studentId)),[myStudents,works]);
  const[editWId,setEditWId]=useState<string|null>(null); const[scores,setScores]=useState<number[]>(Array(10).fill(5));
  function saveEv(){if(!editWId)return;const ex=evals.find(e=>e.workId===editWId);if(ex){setEvals(p=>p.map(e=>e.workId===editWId?{...e,scores,at:new Date().toISOString()}:e))}else{setEvals(p=>[...p,{id:mid("e"),workId:editWId,scores,at:new Date().toISOString(),approved:false}])}setEditWId(null)}
  void saveEv; // called via EvalModal
  return(
    <div className="relative min-h-screen bg-[#0d1230] text-white"><Bg/>
      <header className="sticky top-0 z-20 border-b border-white/10 bg-[#0f1535]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center"><div><div className="text-[10px] font-bold uppercase tracking-[0.28em] text-sky-300">Teacher Evaluation</div><h1 className="text-xl font-black">{teacher.fullName}</h1></div><button onClick={onBack} className="rounded-full border border-white/10 bg-white/5 px-5 py-2.5 text-sm font-bold text-white hover:bg-white/10">Back</button></header>
      <main className="max-w-4xl mx-auto px-6 py-10"><div className="grid gap-4 md:grid-cols-3 mb-10"><Stat l="Students" v={myStudents.length} s="طلابي"/><Stat l="Works" v={myWorks.length} s="الأعمال"/><Stat l="Criteria" v="10" s="عناصر"/></div>
        <div className="grid gap-8 md:grid-cols-2 xl:grid-cols-3">{myWorks.map(w=>{const stu=myStudents.find(s=>s.id===w.studentId);const co=comps.find(c=>c.id===w.competitionId);const ev=evals.find(e=>e.workId===w.id);if(!stu)return null;return(
          <article key={w.id} className="rounded-3xl border border-white/10 bg-white/5 overflow-hidden"><video src={w.videoUrl} poster={w.thumb} controls className="w-full aspect-video object-cover"/><div className="p-5"><div className="text-[10px] font-black uppercase text-sky-300">{co?.nameEn}</div><h3 className="text-base font-black text-white mt-1">{w.title}</h3><p className="text-sm text-[#FDB51F] mt-1">👤 {stu.fullName}</p><button onClick={()=>{const ex=evals.find(e=>e.workId===w.id);setScores(ex?[...ex.scores]:Array(10).fill(5));setEditWId(w.id)}} className={cn("w-full mt-4 rounded-xl px-4 py-3 text-sm font-black transition",ev?"bg-emerald-500/15 border border-emerald-400/20 text-emerald-200":"bg-sky-500 text-white")}>{ev?`Edit Score (${tot(ev.scores)}/100)`:"Add Evaluation ⭐"}</button></div></article>);})}</div></main>
      {editWId&&(<EvalModal workId={editWId} evals={evals} setEvals={setEvals} onClose={()=>setEditWId(null)} teacherName={teacher.fullName} criteria={(() => { const wk=works.find(x=>x.id===editWId); const comp=comps.find(c=>c.id===wk?.competitionId); return comp?.criteria?.length ? comp.criteria : DEFAULT_CRITERIA; })()}/>)}
    </div>
  );
}

/* ════════════════════════════════════
   9. STUDENT PAGE + CONSENT
   ════════════════════════════════════ */
function ConsentModal({student,compAr,compEn,onOk,onNo}:{student:User;compAr:string;compEn:string;onOk:()=>void;onNo:()=>void}){
  const[ok,setOk]=useState(false);
  return(
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/85 p-4 backdrop-blur-md">
      <div className="w-full max-w-2xl rounded-[2.5rem] border border-white/10 bg-[#0a1128] shadow-2xl overflow-hidden max-h-[95vh] flex flex-col"><div className="h-1.5 w-full bg-gradient-to-r from-[#FDB51F] via-[#1697E8] to-[#10B09B] shrink-0"/>
        <div className="p-8 overflow-y-auto"><div className="text-center mb-6"><div className="text-4xl mb-3">📋</div><h2 className="text-2xl font-black text-white">Participation Consent</h2><p className="text-sm text-white/50">إقرار وموافقة المشارك</p></div>
          <div className="rounded-2xl border border-white/10 bg-white/5 p-4 mb-5 flex items-center gap-4"><span className="text-2xl">🎓</span><div><p className="font-black text-white">{student.fullName}</p><p className="text-sm text-white/50">{student.grade} • Section {student.section}</p></div></div>
          <div className="rounded-2xl border border-white/10 bg-white/5 p-5 space-y-5"><div className="border-l-4 border-[#1697E8] pl-4"><p className="text-sm font-black text-[#1697E8] mb-2">📌 Declaration in English</p><p className="text-sm leading-7 text-white/70">I, <b className="text-white">{student.fullName}</b>, confirm my voluntary participation in <b className="text-[#1697E8]">"{compEn}"</b>.</p><ul className="mt-2 space-y-1 text-sm leading-7 text-white/55"><li>✅ My video will be evaluated fairly by a panel of teachers.</li><li>✅ My work will appear on the Champions Celebration Page.</li><li>✅ The submitted work is my own original creation.</li></ul></div>
            <div className="border-t border-white/10"/><div dir="rtl" className="border-r-4 border-[#FDB51F] pr-4"><p className="text-sm font-black text-[#FDB51F] mb-2">📌 إقرار بالعربية</p><p className="text-sm leading-8 text-white/80">أنا <b className="text-white">{student.fullName}</b>، أُقر بمشاركتي الطوعية في مسابقة <b className="text-[#FDB51F]">«{compAr}»</b>.</p><ul className="mt-2 space-y-1 text-sm leading-8 text-white/70"><li>✅ أعلم أن الفيديو سيُعرض للتقييم العادل من لجنة المعلمين.</li><li>✅ أعلم أن مشاركتي ستظهر في صفحة الاحتفاء بالأبطال.</li><li>✅ أؤكد أن العمل المقدم هو عملي الأصلي.</li></ul></div></div>
          <label className="mt-5 flex cursor-pointer items-start gap-4 rounded-2xl border border-[#FDB51F]/20 bg-[#FDB51F]/8 p-4 hover:bg-[#FDB51F]/12"><input type="checkbox" checked={ok} onChange={e=>setOk(e.target.checked)} className="mt-1 h-5 w-5 accent-[#FDB51F]"/><div><p className="text-sm font-black text-[#FFE29A]">I confirm and agree to all terms above.</p><p className="text-sm text-white/40 mt-0.5" dir="rtl">أؤكد رغبتي في المشاركة وأوافق على جميع البنود.</p></div></label>
          <div className="mt-5 flex gap-3"><button onClick={onNo} className="flex-1 rounded-2xl border border-white/10 bg-white/5 px-5 py-4 text-sm font-black text-white/70 hover:bg-white/10">Decline</button><button disabled={!ok} onClick={onOk} className={cn("flex-[2] rounded-2xl px-5 py-4 text-sm font-black shadow-lg",ok?"bg-[#FDB51F] text-slate-950 hover:bg-[#F7C84A]":"cursor-not-allowed bg-white/10 text-white/30")}>✅ Confirm & Upload</button></div>
        </div>
      </div>
    </div>
  );
}

function StudentPage({student,comps,works,evals,certs,selCompId,setWorks,onBack}:{student:User;comps:Comp[];works:Work[];evals:Ev[];certs:Cert[];selCompId:string|null;setWorks:React.Dispatch<React.SetStateAction<Work[]>>;onBack:()=>void}){
  const avail=useMemo(()=>comps.filter(c=>c.started&&!c.archived&&c.grades.includes(student.grade)),[comps,student.grade]);
  const myWorks=useMemo(()=>works.filter(w=>w.studentId===student.id),[student.id,works]);
  const myCerts=useMemo(()=>certs.filter(c=>c.studentId===student.id),[certs,student.id]);
  const[sTab,setSTab]=useState<"participate"|"submissions"|"certs">("participate");
  const[step,setStep]=useState<"idle"|"consent"|"form">("idle");
  const[cid,setCid]=useState(selCompId??avail[0]?.id??"");
  const[title,setTitle]=useState(""); const[vidFile,setVidFile]=useState<File|null>(null); const[vidUrl,setVidUrl]=useState("");
  const[prog,setProg]=useState(0); const[upl,setUpl]=useState(false); const[msg,setMsg]=useState("");
  useEffect(()=>{if(selCompId)setCid(selCompId)},[selCompId]);
  async function submit(e:React.FormEvent){e.preventDefault();if(!cid||!title.trim()||!vidFile){setMsg("All fields required");return}
    setUpl(true);setProg(0);await new Promise<void>(res=>{let p=0;const iv=setInterval(()=>{p+=Math.random()*20+5;if(p>=100){clearInterval(iv);setProg(100);setUpl(false);res()}else setProg(Math.round(p))},200)});
    setWorks(p=>[{id:mid("w"),competitionId:cid,studentId:student.id,title:title.trim(),desc:"",videoUrl:vidUrl,thumb:vidUrl,at:new Date().toISOString()},...p]);
    setMsg("🎉 Submitted! Your work will appear on the Champions page.");setTitle("");setVidFile(null);setVidUrl("");setStep("idle")}
  const selComp=avail.find(c=>c.id===cid);
  return(
    <div className="relative min-h-screen bg-[#0d1230] text-white"><Bg/>
      {step==="consent"&&selComp&&<ConsentModal student={student} compAr={selComp.nameAr} compEn={selComp.nameEn} onOk={()=>setStep("form")} onNo={()=>setStep("idle")}/>}
      <header className="sticky top-0 z-20 border-b border-white/10 bg-[#0f1535]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center"><div><div className="text-[10px] font-bold uppercase tracking-[0.28em] text-[#FDB51F]">Student Portal</div><h1 className="text-xl font-black">{student.fullName}</h1><p className="text-sm text-white/50">{student.grade} • {student.section}</p></div><button onClick={onBack} className="rounded-full border border-white/10 bg-white/5 px-5 py-2.5 text-sm font-bold text-white hover:bg-white/10">Back</button></header>
      <main className="max-w-6xl mx-auto px-6 py-8">
        <div className="flex flex-wrap gap-2 mb-8 bg-white/5 border border-white/10 rounded-2xl p-1.5">
          <button onClick={()=>setSTab("participate")} className={cn("rounded-xl px-5 py-3 text-sm font-black",sTab==="participate"?"bg-[#FDB51F] text-slate-950":"text-white/50")}>🎓 Participate</button>
          <button onClick={()=>setSTab("submissions")} className={cn("rounded-xl px-5 py-3 text-sm font-black",sTab==="submissions"?"bg-[#1697E8] text-white":"text-white/50")}>📁 My Works ({myWorks.length})</button>
          <button onClick={()=>setSTab("certs")} className={cn("rounded-xl px-5 py-3 text-sm font-black",sTab==="certs"?"bg-emerald-500 text-white":"text-white/50")}>📜 Certificates ({myCerts.length})</button>
        </div>
        {sTab==="participate"&&(
          <div className="max-w-xl"><div className="rounded-[2rem] border border-white/10 bg-white/5 p-6 backdrop-blur-sm">
            {step==="idle"&&(<><h2 className="text-xl font-black mb-4">Participate</h2><div className="space-y-4"><div><FL>Select Competition</FL><Sel value={cid} onChange={e=>setCid(e.target.value)}><option value="">— Select —</option>{avail.map(c=><option key={c.id} value={c.id}>{c.nameEn}</option>)}</Sel></div>{msg&&<div className="p-3 bg-emerald-500/10 border border-emerald-400/20 text-emerald-200 rounded-xl text-sm">{msg}</div>}<button disabled={!cid} onClick={()=>{setMsg("");setStep("consent")}} className={cn("w-full rounded-2xl px-5 py-4 text-sm font-black",cid?"bg-[#FDB51F] text-slate-950 hover:bg-[#F7C84A] shadow-lg":"cursor-not-allowed bg-white/10 text-white/30")}>📋 Read Consent & Participate</button></div></>)}
            {step==="form"&&(<form onSubmit={submit} className="space-y-5"><div className="flex items-center gap-3 mb-2"><button type="button" onClick={()=>setStep("idle")} className="text-white/50 text-xs">← Back</button><p className="text-xs text-emerald-300 font-bold">✅ Consent Accepted</p></div><div><FL>Title</FL><Inp value={title} onChange={e=>setTitle(e.target.value)} placeholder="Work title"/></div><div><FL>📹 Upload Video</FL><label className={cn("flex flex-col items-center gap-3 rounded-2xl border-2 border-dashed p-6 cursor-pointer transition",vidFile?"border-emerald-400/40 bg-emerald-500/5":"border-white/10 bg-white/5 hover:border-white/25")}><input type="file" accept="video/*" onChange={e=>{const f=e.target.files?.[0];if(f){setVidFile(f);setVidUrl(URL.createObjectURL(f))}}} className="hidden"/>{vidFile?(<><span className="text-3xl">🎬</span><p className="font-black text-emerald-300">{vidFile.name}</p><video src={vidUrl} controls className="w-full max-h-28 rounded-xl mt-2"/></>):(<><span className="text-3xl">📹</span><p className="text-sm font-black text-white">Click to select</p><p className="text-xs text-white/40">MP4, MOV, AVI</p></>)}</label></div>{upl&&<div className="rounded-2xl border border-sky-400/20 bg-sky-500/10 p-4"><div className="flex justify-between text-xs font-bold text-sky-200 mb-2"><span>Uploading…</span><span>{prog}%</span></div><div className="h-2 w-full rounded-full bg-white/10 overflow-hidden"><div className="h-2 rounded-full bg-gradient-to-r from-[#1697E8] to-[#10B09B]" style={{width:`${prog}%`}}/></div></div>}<button type="submit" disabled={upl} className={cn("w-full rounded-2xl px-5 py-4 text-sm font-black shadow-lg",upl?"cursor-not-allowed bg-white/10 text-white/30":"bg-[#FDB51F] text-slate-950 hover:bg-[#F7C84A]")}>{upl?`${prog}%...`:"🚀 Submit Work"}</button></form>)}
          </div></div>
        )}
        {sTab==="submissions"&&(<div className="grid gap-6 md:grid-cols-2 xl:grid-cols-3">{myWorks.length===0&&<div className="md:col-span-3 rounded-3xl border border-white/10 bg-white/5 p-12 text-center text-white/40">No submissions yet</div>}{myWorks.map(w=>{const ev=evals.find(e=>e.workId===w.id);const co=comps.find(c=>c.id===w.competitionId);return(<article key={w.id} className="rounded-2xl border border-white/10 bg-white/5 overflow-hidden"><video src={w.videoUrl} poster={w.thumb} controls className="w-full aspect-video object-cover"/><div className="p-4"><div className="text-[10px] font-black uppercase text-sky-300">{co?.nameEn}</div><h3 className="text-base font-black text-white mt-1">{w.title}</h3><div className="mt-3 flex flex-wrap gap-2">{ev?<Pill t="green">✅ Score: {tot(ev.scores)}/100</Pill>:<Pill t="slate">⏳ Under Review</Pill>}</div></div></article>);})}</div>)}
        {sTab==="certs"&&(<div className="space-y-6">{myCerts.length===0&&<div className="rounded-3xl border border-white/10 bg-white/5 p-12 text-center"><p className="text-4xl mb-3">📜</p><p className="text-white/40">No certificates yet. Certificates are issued automatically when a competition closes.</p></div>}{myCerts.map(cert=><article key={cert.id} className="rounded-3xl border border-[#FDB51F]/30 bg-gradient-to-r from-[#FDB51F]/10 to-[#0d1230] p-6 flex items-center gap-5 justify-between flex-wrap"><div className="flex items-center gap-4"><div className="text-4xl">{MEDALS[cert.rank-1]||"🏅"}</div><div><h3 className="text-xl font-black text-white">{cert.compName}</h3><div className="flex gap-2 mt-2"><Pill t="gold">Rank #{cert.rank}</Pill><Pill t="green">Score: {cert.score}/100</Pill><Pill t="blue">{cert.grade}</Pill></div></div></div></article>)}</div>)}
      </main>
    </div>
  );
}
