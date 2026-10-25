import React, { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import QRCode from "qrcode";
import { supabase, isSupabaseConfigured } from "./lib/supabase";
import { uploadVideo, deleteVideo } from "./lib/videoUpload";

/* ═══════ TYPES ═══════ */
type Role = "teacher" | "student";
type View = "home"|"showcase"|"classBoard"|"allWorks"|"winners"|"adminLogin"|"admin"|"teacherLogin"|"teacher"|"teacherPortal"|"studentLogin"|"student";
type User = { id:string; fullName:string; role:Role; grade:string; section:string; teacherId?:string; username:string; password:string; createdAt:string; active:boolean };
type Comp = { id:string; nameAr:string; nameEn:string; logoUrl:string; grades:string[]; rulesAr:string; rulesEn:string; closingDate:string; certDate:string; started:boolean; archived:boolean; criteria:string[]; evalCode:string };
type WorkStatus = "pending" | "approved";
type Work = { id:string; competitionId:string; studentId:string; title:string; desc:string; videoUrl:string; thumb:string; at:string; status?:WorkStatus; approvedBy?:string; approvedAt?:string };
type Ev = { id:string; workId:string; scores:number[]; at:string; approved?:boolean };
type Like = { workId:string; type:"friend"|"parent"|"teacher"; name:string; at:string };
type Cert = { id:string; workId:string; studentId:string; compId:string; compName:string; score:number; rank:number; grade:string; section:string; issuedAt:string };
type ClassInfo = { grade:string; section:string; teacherId:string; teacherName:string; count:number };
type ArchiveEntry = { id:string; compId:string; compName:string; compNameAr:string; archivedAt:string; winners:Array<{rank:number; studentName:string; grade:string; section:string; score:number}> };

/* ═══════ CONSTANTS ═══════ */
const SK = { u:"gwa-u-v8", c:"gwa-c-v8", w:"gwa-w-v8", e:"gwa-e-v8", l:"gwa-l-v8", cert:"gwa-cert-v8", arch:"gwa-arch-v8", pw:"gwa-pw-v8" };
const GRADES = ["KG 1","KG 2","Grade 1","Grade 2","Grade 3","Grade 4","Grade 5","Grade 6","Grade 7","Grade 8","Grade 9","Grade 10","Grade 11","Grade 12"];
const _SECTIONS = ["A","B","C","D","X","Y"];
const DEFAULT_CRITERIA = ["Creativity • الإبداع","Clarity • الوضوح","Content • المحتوى","Language • اللغة","Confidence • الثقة","Design • التصميم","Compliance • الالتزام","Organization • التنظيم","Technology • التقنية","Final Impact • الإخراج"];
const MEDALS = ["🥇","🥈","🥉","4️⃣","5️⃣","6️⃣","7️⃣","8️⃣","9️⃣","🔟"];

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


/* ═══════ STANDARD COMPETITION RULES (fixed template, editable by admin) ═══════ */
const RULE_LIMITS_AR = "• مدة الفيديو: 3 دقائق كحد أقصى.\n• الحد الأقصى لحجم الفيديو: 50 ميجابايت (حد التخزين السحابي الحالي).";
const RULE_LIMITS_EN = "• Maximum video duration: 3 minutes.\n• Maximum video file size: 50 MB (current cloud storage limit).";

function ruleTemplateAr(nameAr:string){
  return `قواعد مسابقة «${nameAr}»\n`
    + "١- الالتزام بموضوع المسابقة والضوابط المعلنة.\n"
    + "٢- أن يكون العمل من إنتاج الطالب نفسه.\n"
    + "٣- تصوير الفيديو أفقيًا وبإضاءة وصوت واضحين.\n"
    + "٤- ذكر اسم الطالب وصفه في بداية الفيديو.\n"
    + "٥- أن يكون الفيديو بصيغة MP4 أو MOV.\n"
    + "٦- لا يُعرض الفيديو للجميع إلا بعد اعتماد أحد المعلمين.\n"
    + "٧- لا يُقبل أي عمل بعد تاريخ إغلاق المسابقة.\n"
    + RULE_LIMITS_AR;
}
function ruleTemplateEn(nameEn:string){
  return `Rules for "${nameEn}"\n`
    + "1- Follow the competition topic and the published guidelines.\n"
    + "2- The submitted work must be the student's own.\n"
    + "3- Record the video in landscape with clear light and sound.\n"
    + "4- State your full name and grade at the start of the video.\n"
    + "5- Accepted formats: MP4 or MOV.\n"
    + "6- A submission stays locked until a teacher approves it.\n"
    + "7- No submissions are accepted after the closing date.\n"
    + RULE_LIMITS_EN;
}
function rulesForPreset(idx?:number, nameAr?:string, nameEn?:string){
  const preset = (idx!==undefined && !isNaN(idx)) ? COMP_PRESETS[idx] : undefined;
  const ar = preset?.nameAr ?? nameAr ?? "المسابقة";
  const en = preset?.nameEn ?? nameEn ?? "Competition";
  return { ar: ruleTemplateAr(ar), en: ruleTemplateEn(en) };
}

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

const seedU: User[] = [
  /* ONE default teacher account — teaches several classes */
  { id:"t1", fullName:"Ms Heba", role:"teacher", grade:"Grade 6:X;Grade 5:A", section:"", username:"heba", password:"Teach2050", createdAt:"", active:true },
  /* ONE default student account */
  { id:"s1", fullName:"Mariam Salem", role:"student", grade:"Grade 6", section:"X", teacherId:"t1", username:"mariam.salem", password:"GWA2050", createdAt:"", active:true },
];

const seedC:Comp[]=[
  {id:"cq1",nameEn:"Holy Quran Recitation Competition",nameAr:"مسابقة تلاوة القرآن الكريم",logoUrl:"/gems-logo.png",grades:["Grade 4","Grade 5","Grade 6","Grade 7","Grade 8"],rulesAr:"١- يتلو الطالب آيات من القرآن الكريم\n٢- يُقيّم على أساس التجويد والأداء\n٣- مدة الفيديو ٣ دقائق\n٤- الالتزام بأحكام التلاوة",rulesEn:"1- Student recites verses from the Holy Quran\n2- Evaluated on Tajweed and delivery\n3- Video duration: 3 minutes\n4- Adherence to recitation rules",closingDate:"2027-02-01",certDate:"2027-02-15",started:true,archived:false,criteria:COMP_PRESETS[0].criteria,evalCode:"QURAN1"},
  {id:"ct1",nameEn:"Arabic Big Talk",nameAr:"مسابقة التحدث بالعربية",logoUrl:"/gems-logo.png",grades:["Grade 4","Grade 5","Grade 6","Grade 7","Grade 8"],rulesAr:"١- يتحدث الطالب بالعربية الفصحى لمدة ٢-٣ دقائق\n٢- يُقيّم على الطلاقة والنطق ولغة الجسد\n٣- الموضوع حر\n٤- مفتوحة للطلاب العرب وغير العرب",rulesEn:"1- Student speaks in Arabic for 2-3 minutes\n2- Evaluated on fluency, pronunciation & body language\n3- Free topic choice\n4- Open to Arab and non-Arab students",closingDate:"2027-02-01",certDate:"2027-02-15",started:true,archived:false,criteria:COMP_PRESETS[1].criteria,evalCode:"TALK1"},
  {id:"cp1",nameEn:"Classical Arabic Poetry Recitation",nameAr:"مسابقة الشعر العربي الأصيل",logoUrl:"/gems-logo.png",grades:["Grade 5","Grade 6","Grade 7","Grade 8","Grade 9"],rulesAr:"١- يُلقي الطالب قصيدة من الشعر العربي الكلاسيكي\n٢- التقييم على الأداء والإيقاع والتعبير\n٣- مدة الإلقاء ٢-٤ دقائق\n٤- يُفضل الحفظ الكامل",rulesEn:"1- Student recites a classical Arabic poem\n2- Evaluated on delivery, rhythm & expression\n3- Duration: 2-4 minutes\n4- Full memorization preferred",closingDate:"2027-02-01",certDate:"2027-02-15",started:true,archived:false,criteria:COMP_PRESETS[3].criteria,evalCode:"POETRY1"},
  {id:"cr1",nameEn:"Arabic Loud Reading Competition",nameAr:"مسابقة القراءة الجهرية",logoUrl:"/gems-logo.png",grades:["Grade 3","Grade 4","Grade 5","Grade 6","Grade 7"],rulesAr:"١- يقرأ الطالب نصاً عربياً قراءة جهرية واضحة\n٢- التقييم على النطق والطلاقة والفهم\n٣- مدة القراءة ٢-٣ دقائق\n٤- النص يُختار من الأدب العربي",rulesEn:"1- Student reads an Arabic text aloud clearly\n2- Evaluated on pronunciation, fluency & comprehension\n3- Duration: 2-3 minutes\n4- Text selected from Arabic literature",closingDate:"2027-02-01",certDate:"2027-02-15",started:true,archived:false,criteria:COMP_PRESETS[4].criteria,evalCode:"READ1"},
  {id:"cu1",nameEn:"UAE: Land of Goodness",nameAr:"مسابقة الاجتماعيات: الإمارات وطن الخير",logoUrl:"/gems-logo.png",grades:["Grade 4","Grade 5","Grade 6","Grade 7","Grade 8"],rulesAr:"١- يقدم الطالب مشروعاً عن الإمارات العربية المتحدة\n٢- يمكن أن يكون فيديو أو عرض رقمي\n٣- يركز على التاريخ أو الجغرافيا أو القيم الوطنية\n٤- الأصالة والإبداع أساس التقييم",rulesEn:"1- Student presents a project about the UAE\n2- Can be a video or digital presentation\n3- Focus on history, geography or national values\n4- Originality and creativity are key",closingDate:"2027-02-01",certDate:"2027-02-15",started:true,archived:false,criteria:COMP_PRESETS[8].criteria,evalCode:"UAE1"},
];

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
function sv(k:string,v:any){
  try{ localStorage.setItem(k,JSON.stringify(v)); }catch{}
  if(typeof window === 'undefined' || !isSupabaseConfigured) return;
  const validDate=(d:any)=>{ if(!d) return new Date().toISOString(); const t=Date.parse(d); return isNaN(t)?new Date().toISOString():new Date(t).toISOString(); };
  const chunk=<T,>(arr:T[],size:number):T[][]=>{const out:T[][]=[];for(let i=0;i<arr.length;i+=size)out.push(arr.slice(i,i+size));return out};
  setTimeout(async ()=>{
    try {
      /* ── ACCOUNTS: dedupe by id + username so one conflict never breaks the whole sync ── */
      if(k===SK.u && Array.isArray(v) && v.length > 0){
        const byId = new Map<string, any>();
        v.forEach((u:any)=>{
          if(!u || !u.id || !u.username) return;
          byId.set(String(u.id), {
            id:String(u.id), full_name:u.fullName||u.username, role:u.role||"student",
            grade:u.grade||"", section:u.section||"", teacher_id:u.teacherId||null,
            username:String(u.username).trim().toLowerCase(), password:String(u.password ?? ""),
            active:u.active!==false
          });
        });
        const byName = new Map<string, any>();
        Array.from(byId.values()).forEach(r=>{ byName.set(r.username, r); });
        const rows = Array.from(byName.values());
        for(const batch of chunk(rows, 40)){
          const {error} = await supabase.from("accounts").upsert(batch,{onConflict:"id"});
          if(error){
            console.error("accounts upsert:", error.message);
            for(const row of batch){
              const {error:rowErr} = await supabase.from("accounts").upsert(row,{onConflict:"id"});
              if(rowErr) console.error("account row skipped:", row.username, rowErr.message);
            }
          }
        }
      }
      /* ── COMPETITIONS: mirror to the cloud so edits/deletes reach every device ── */
      if(k===SK.c && Array.isArray(v)){
        const rows = v.map((c:any)=>({
          id:String(c.id), name_en:c.nameEn||"", name_ar:c.nameAr||"",
          logo_url:c.logoUrl||"/gems-logo.png",
          grades:Array.isArray(c.grades)?c.grades:[],
          rules_ar:c.rulesAr||"", rules_en:c.rulesEn||"",
          closing_date:c.closingDate||"", cert_date:c.certDate||"",
          started:c.started!==false, archived:c.archived===true,
          criteria:Array.isArray(c.criteria)?c.criteria:[],
          eval_code:c.evalCode||""
        }));
        const sentinel = {
          id:"__meta__", name_en:"__meta__", name_ar:"__meta__", logo_url:"",
          grades:[], rules_ar:"", rules_en:"", closing_date:"", cert_date:"",
          started:true, archived:true, criteria:[], eval_code:""
        };
        const payload = [sentinel, ...rows];
        if(payload.length>0){
          const {error} = await supabase.from("competitions_cloud").upsert(payload,{onConflict:"id"});
          if(error) console.warn("competitions sync skipped:", error.message);
        }
        const {data:cloudIds,error:listErr} = await supabase.from("competitions_cloud").select("id");
        if(!listErr && Array.isArray(cloudIds)){
          const keep = new Set([...rows.map((r:any)=>r.id),"__meta__"]);
          const stale = cloudIds.map((r:any)=>String(r.id)).filter((id:string)=>!keep.has(id));
          for(const id of stale){
            await supabase.from("competitions_cloud").delete().eq("id",id);
          }
        }
      }
      if(k===SK.w && Array.isArray(v) && v.length > 0){
        const rows = v.map((w:any)=>({
          id:String(w.id), competition_id:String(w.competitionId||""), student_id:String(w.studentId||""),
          title:w.title||"Untitled", desc:w.desc||"", video_url:w.videoUrl||"", thumb:w.thumb||"",
          at: validDate(w.at),
          status:w.status==="approved"?"approved":"pending",
          approved_by:w.approvedBy||null,
          approved_at:w.approvedAt?validDate(w.approvedAt):null
        }));
        const {error} = await supabase.from("works_cloud").upsert(rows,{onConflict:"id"});
        if(error){
          console.error("works upsert:", error.message);
          // Older cloud table without the approval columns → still save the submission
          const legacy = rows.map((r:any)=>{ const {status,approved_by,approved_at,...rest}=r; return rest; });
          const {error:legacyErr} = await supabase.from("works_cloud").upsert(legacy,{onConflict:"id"});
          if(legacyErr) console.error("works upsert (legacy):", legacyErr.message);
        }
      }
      if(k===SK.e && Array.isArray(v) && v.length > 0){
        const rows = v.map((e:any)=>({
          id:String(e.id), work_id:String(e.workId||""), scores:e.scores||[], approved:e.approved||false, at: validDate(e.at)
        }));
        let {error} = await supabase.from("evaluations").upsert(rows,{onConflict:"id"});
        if(error){
          await new Promise(r=>setTimeout(r,1500));
          const retry = await supabase.from("evaluations").upsert(rows,{onConflict:"id"});
          error = retry.error;
        }
        if(error) console.error("evaluations upsert:", error.message);
      }
      if(k===SK.cert && Array.isArray(v) && v.length > 0){
        const rows = v.map((c:any)=>({
          id:String(c.id), work_id:String(c.workId||""), student_id:String(c.studentId||""), comp_id:String(c.compId||""),
          comp_name:c.compName||"", score:Number(c.score)||0, rank:Number(c.rank)||0, grade:c.grade||"",
          section:c.section||"", issued_at:c.issuedAt||""
        }));
        const {error} = await supabase.from("certificates").upsert(rows,{onConflict:"id"});
        if(error) console.error("certificates upsert:", error.message);
      }
    } catch(err){ console.error("Supabase save error:", err); }
  }, 120);
}
function tot(s:number[]){return s.reduce((a,b)=>a+b,0)}
function slug(n:string){return n.trim().toLowerCase().replace(/[^a-z0-9\s]/g,"").replace(/\s+/g,".").replace(/^\.+|\.+$/g,"")}
function mkUser(name:string,role:Role,grade:string,section:string,tid:string|undefined,users:User[]):User{
  let u=slug(name)||"gwa.user";let c=1;while(users.some(x=>x.username===u)){u=slug(name)+c;c++}
  const num=String(users.length+2050).slice(-4);
  return{id:mid(role.charAt(0)),fullName:name.trim(),role,grade,section:role==="student"?section:"",teacherId:tid,username:u,password:(role==="teacher"?"Teach":"GWA")+num,createdAt:new Date().toISOString(),active:true}
}


/* ═══════ TEACHER CLASSES (a teacher may teach several classes) ═══════ */
type ClassRef = { grade:string; section:string };
const CLASS_SEP = ";";
const CLASS_PART = ":";
function encodeClasses(list:ClassRef[]):string{
  const seen=new Set<string>();
  return list
    .filter(c=>c && c.grade)
    .filter(c=>{const k=`${c.grade}${CLASS_PART}${c.section||""}`;if(seen.has(k))return false;seen.add(k);return true;})
    .map(c=>`${c.grade}${CLASS_PART}${c.section||""}`)
    .join(CLASS_SEP);
}
function decodeClasses(u:{role?:string;grade?:string;section?:string}|null|undefined):ClassRef[]{
  if(!u||!u.grade) return [];
  if(u.role!=="teacher") return [{grade:u.grade, section:u.section||""}];
  const out:ClassRef[]=[];
  String(u.grade).split(CLASS_SEP).forEach(part=>{
    const raw=String(part||"").trim();
    if(!raw) return;
    if(raw.includes(CLASS_PART)){
      const [g,sec]=raw.split(CLASS_PART);
      out.push({grade:(g||"").trim(), section:(sec||"").trim()});
    }else{
      out.push({grade:raw, section:(u.section||"").trim()});
    }
  });
  if(out.length===0) return [{grade:u.grade, section:u.section||""}];
  return out;
}
function classLabel(c:ClassRef){ return `${c.grade}${c.section?` • ${c.section}`:""}`; }
function classesText(u:any){ const list=decodeClasses(u); return list.length?list.map(classLabel).join(" , "):"—"; }

/* ═══════ ADMIN RECOVERY ═══════ */
const ADMIN_RECOVERY_CODE = "MohdRoshdi@2050";
const ADMIN_EMAIL = "m.roshdi_gwa@gemsedu.com";        /* primary • البريد الأساسي */
const ADMIN_EMAIL_ALT = "mohdarabichod3@gmail.com";   /* alternate • بريد بديل مقبول */
const ADMIN_EMAILS = [ADMIN_EMAIL, ADMIN_EMAIL_ALT];
const isAdminEmail = (value:string)=> ADMIN_EMAILS.includes(String(value||"").trim().toLowerCase());

/* ═══════ SUPABASE HELPERS ═══════ */
function useVideoUrl(path: string): string {
  const [url, setUrl] = useState("");
  useEffect(() => {
    if (!path) return;
    if (path.startsWith("http") || path.startsWith("/")) { setUrl(path); return; }
    if (!isSupabaseConfigured) return;
    
    // Call Edge Function to get signed URL
    supabase.functions.invoke("get-video-url", { body: { path } })
      .then(({ data, error }) => {
        if (!error && data?.signedUrl) setUrl(data.signedUrl);
      });
  }, [path]);
  return url;
}

function SecureVideo(props: React.VideoHTMLAttributes<HTMLVideoElement> & { path: string }) {
  const { path, ...videoProps } = props;
  const url = useVideoUrl(path);
  
  if (!url) {
    return (
      <div className={cn("flex items-center justify-center bg-slate-900 animate-pulse", props.className)}>
        <span className="text-xs text-white/20 font-bold uppercase tracking-widest">Loading Secure Video...</span>
      </div>
    );
  }

  return <video {...videoProps} src={url} />;
}


/* ═══════ REVIEW / APPROVAL RULES ═══════ */
const MAX_VIDEO_MB = 50;
const MAX_VIDEO_SECONDS = 180; // 3 minutes — matches the current Supabase storage tier
const isApproved = (w:Work)=> w.status === "approved";

function StatusPill({w}:{w:Work}){
  if(isApproved(w)){
    return <span className="inline-flex items-center gap-1.5 rounded-full border border-emerald-400/30 bg-emerald-500/15 px-2.5 py-1 text-[10px] font-black text-emerald-200">✅ Approved • معتمد{w.approvedBy?` · ${w.approvedBy}`:""}</span>;
  }
  return <span className="inline-flex items-center gap-1.5 rounded-full border border-[#FDB51F]/40 bg-[#FDB51F]/15 px-2.5 py-1 text-[10px] font-black text-[#FFE29A]">⏳ Waiting Teachers Approval 1st</span>;
}

/* Teacher approves a submission by typing their own username */
function TeacherApproveModal({work,users,setWorks,onClose}:{work:Work;users:User[];setWorks:React.Dispatch<React.SetStateAction<Work[]>>;onClose:()=>void}){
  const[username,setUsername]=useState(""); const[password,setPassword]=useState("");
  const[err,setErr]=useState(""); const[busy,setBusy]=useState(false);
  const norm=(x:any)=>String(x||"").trim().toLowerCase();
  const mapRow=(r:any):User=>({id:r.id,fullName:r.full_name||r.username||"Teacher",role:"teacher",grade:r.grade||"",section:r.section||"",teacherId:undefined,username:r.username,password:String(r.password??""),createdAt:r.created_at||"",active:r.active!==false});

  async function submit(e:React.FormEvent){
    e.preventDefault();
    const uname=username.trim(); const pass=password;
    if(!uname||!pass){setErr("أدخل اسم المستخدم وكلمة المرور • Enter username and password");return}
    setBusy(true); setErr("");
    try{
      /* 1) cached/synced list first */
      let teacher=users.find(x=>x.role==="teacher" && norm(x.username)===norm(uname) && String(x.password)===pass && x.active!==false);
      /* 2) cloud is the source of truth — works on any device */
      if(!teacher && isSupabaseConfigured){
        const{data,error}=await supabase.from("accounts").select("*").ilike("username",uname).limit(20);
        if(error) console.warn("approval lookup:",error.message);
        const rows=((data||[]) as any[]);
        const matched=rows.find(r=>String(r.password??"")===pass && r.role==="teacher" && r.active!==false);
        if(matched) teacher=mapRow(matched);
        else if(rows.some(r=>norm(r.username)===norm(uname))){
          setErr("كلمة المرور غير صحيحة • Incorrect password"); setBusy(false); return;
        }
      }
      if(!teacher){setErr("اسم المستخدم غير موجود أو ليس حساب معلم • Username not found or not a teacher account");setBusy(false);return}
      if(teacher.id===work.studentId){setErr("لا يمكن للمعلم اعتماد فيديو خاص به • A teacher cannot approve their own upload");setBusy(false);return}
      setWorks(p=>p.map(x=>x.id===work.id?{...x,status:"approved" as WorkStatus,approvedBy:teacher!.fullName,approvedAt:new Date().toISOString()}:x));
      onClose();
    }catch(e:any){setErr("تعذر الاعتماد • "+(e?.message||"error"))}
    finally{setBusy(false)}
  }

  return(
    <div className="fixed inset-0 z-[70] flex items-center justify-center bg-black/85 p-4 backdrop-blur-md" onClick={onClose}>
      <div className="w-full max-w-md overflow-hidden rounded-[2rem] border border-[#FDB51F]/25 bg-[#0f1535] shadow-2xl" onClick={e=>e.stopPropagation()}>
        <div className="h-1.5 w-full bg-gradient-to-r from-[#FDB51F] via-[#10B09B] to-[#1697E8]"/>
        <div className="p-6">
          <div className="text-center mb-4">
            <div className="mx-auto mb-2 grid h-12 w-12 place-items-center rounded-2xl bg-emerald-500/15 text-2xl ring-1 ring-emerald-400/25">🔓</div>
            <h3 className="text-lg font-black text-white">Approve Submission</h3>
            <p className="text-sm text-white/50">اعتماد الفيديو — أي معلم باسمه وكلمة مروره</p>
          </div>
          <div className="rounded-2xl border border-white/10 bg-white/5 p-3 mb-4">
            <p className="text-sm font-black text-white line-clamp-1">{work.title}</p>
            <p className="text-xs text-white/40 mt-0.5">{users.find(u=>u.id===work.studentId)?.fullName||"—"}</p>
          </div>
          <form onSubmit={submit} className="space-y-3">
            <div><FL>Teacher Username • اسم مستخدم المعلم</FL><Inp value={username} onChange={e=>{setUsername(e.target.value);setErr("")}} placeholder="e.g. heba" autoFocus autoComplete="off"/></div>
            <div><FL>Password • كلمة المرور</FL><Inp type="password" value={password} onChange={e=>{setPassword(e.target.value);setErr("")}} placeholder="••••••••" autoComplete="off"/></div>
            {err&&<div className="rounded-xl border border-rose-400/20 bg-rose-500/10 px-4 py-3 text-sm font-bold leading-6 text-rose-200">{err}</div>}
            <div className="flex gap-3">
              <button type="button" onClick={onClose} className="flex-1 rounded-2xl border border-white/10 bg-white/5 py-3 text-sm font-black text-white/70 hover:bg-white/10">Cancel</button>
              <button type="submit" disabled={busy} className={cn("flex-[1.4] rounded-2xl py-3 text-sm font-black transition",busy?"cursor-not-allowed bg-white/10 text-white/30":"bg-emerald-500 text-slate-950 hover:bg-emerald-400")}>{busy?"Checking…":"✅ Approve • اعتماد"}</button>
            </div>
            <p className="text-center text-[10px] leading-5 text-white/30">Can also be approved from the Teacher Portal • يمكن الاعتماد أيضاً من بوابة المعلم</p>
          </form>
        </div>
      </div>
    </div>
  );
}

function ApprovalGate({work,users,setWorks}:{work:Work;users:User[];setWorks:React.Dispatch<React.SetStateAction<Work[]>>}){
  const[open,setOpen]=useState(false);
  return(<>
    <button onClick={()=>setOpen(true)} className="rounded-full border border-[#FDB51F]/40 bg-[#FDB51F]/15 px-4 py-2 text-[11px] font-black text-[#FFE29A] hover:bg-[#FDB51F]/25 transition">🔓 Teacher approve • اعتماد</button>
    {open&&<TeacherApproveModal work={work} users={users} setWorks={setWorks} onClose={()=>setOpen(false)}/>}
  </>);
}

/* Video area: locked until a teacher approves */
function WorkVideo({w,users,setWorks}:{w:Work;users:User[];setWorks:React.Dispatch<React.SetStateAction<Work[]>>}){
  if(isApproved(w)){
    return <SecureVideo path={w.videoUrl} poster={w.thumb} controls preload="metadata" className="w-full aspect-video object-cover"/>;
  }
  return(
    <div className="relative">
      <div className="w-full aspect-video bg-[#0b1028] grid place-items-center"><span className="text-5xl opacity-25">🔒</span></div>
      <div className="absolute inset-0 flex flex-col items-center justify-center gap-2.5 bg-[#0b1028]/90 p-4 text-center">
        <p className="text-sm font-black text-[#FFE29A]">⏳ Waiting Teachers Approval 1st</p>
        <p className="text-[11px] font-bold leading-5 text-white/50" dir="rtl">الفيديو مغلق لحين معاينته من أحد المعلمين<br/>Locked until a teacher reviews it</p>
        <ApprovalGate work={w} users={users} setWorks={setWorks}/>
      </div>
    </div>
  );
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

/* ── EVALUATION EDITOR MODAL ── */
function EvalModal({workId,evals,setEvals,onClose,teacherName,criteria=DEFAULT_CRITERIA}:{workId:string;evals:Ev[];setEvals:React.Dispatch<React.SetStateAction<Ev[]>>;onClose:()=>void;teacherName:string;criteria:string[]}){
  const ex=evals.find(e=>e.workId===workId);
  const[scores,setScores]=useState<number[]>(ex?[...ex.scores]:Array(10).fill(5));
  function save(approved=false){
    if(ex){
      setEvals(p=>p.map(e=>e.workId===workId?{...e,scores,at:new Date().toISOString(),approved}:e));
    }else{
      setEvals(p=>[...p,{id:mid("e"),workId,scores,at:new Date().toISOString(),approved}]);
    }
    onClose();
  }
  return(
    <div className="fixed inset-0 z-50 bg-black/85 backdrop-blur-md flex items-center justify-center p-4">
      <div className="bg-[#0f1535] border border-white/10 rounded-[2rem] w-full max-w-3xl max-h-[90vh] shadow-2xl flex flex-col overflow-hidden">
        <div className="p-6 border-b border-white/10 flex justify-between items-center shrink-0">
          <div><h2 className="text-2xl font-black text-[#FDB51F]">Evaluation • تحرير التقييم</h2><p className="text-sm text-white/50">Teacher: {teacherName}</p></div>
          <button onClick={onClose} className="rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm font-bold text-white">✕</button>
        </div>
        <div className="flex-1 overflow-y-auto p-6 space-y-4">
          {criteria.map((cr:string,i:number)=>(
            <div key={i} className="bg-white/5 p-4 rounded-2xl border border-white/10">
              <div className="flex justify-between mb-2"><span className="font-bold text-white/80 text-sm">{cr}</span><span className="font-black text-sky-300">{scores[i]||0}/10</span></div>
              <input type="range" min="0" max="10" value={scores[i]||0} onChange={e=>{const n=[...scores];n[i]=+e.target.value;setScores(n)}} className="w-full accent-sky-400 cursor-pointer"/>
            </div>
          ))}
        </div>
        <div className="p-6 border-t border-white/10 bg-[#0f1535] shrink-0 sticky bottom-0">
          <div className="flex flex-col sm:flex-row gap-4 items-center justify-between">
            <div className="rounded-2xl border border-[#FDB51F]/20 bg-[#FDB51F]/10 px-5 py-4 text-sm font-black text-[#FFE29A]">Total: {tot(scores)}/100</div>
            <div className="flex flex-wrap gap-3 justify-end">
              <button onClick={()=>save(false)} className="rounded-2xl border border-white/10 bg-white/5 px-6 py-4 text-sm font-black text-white hover:bg-white/10">Save Draft 💾</button>
              <button onClick={()=>save(true)} className="rounded-2xl bg-[#FDB51F] px-6 py-4 text-sm font-black text-slate-950 shadow-lg hover:bg-[#F7C84A]">Approve Evaluation ✅</button>
            </div>
          </div>
        </div>
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
function WorkCard({w,users,evals,comps,likes,setLikes,setEvals,setWorks,certs,onPromote}:{w:Work;users:User[];evals:Ev[];comps:Comp[];likes:Like[];setLikes:React.Dispatch<React.SetStateAction<Like[]>>;setEvals:React.Dispatch<React.SetStateAction<Ev[]>>;setWorks:React.Dispatch<React.SetStateAction<Work[]>>;certs:Cert[];onPromote:(wid:string,cid:string,code:string)=>boolean}){
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

  async function verifyCode(){
    if(!comp){setCodeErr("No competition found");return}
    if(codeInput.trim().toUpperCase()===comp.evalCode?.toUpperCase()){
      if(codeMode==="eval"){
        setShowCodePrompt(false); setCodeInput(""); setCodeErr("");
        setEditEval(true);
      }else{
        // Delete flow
        try {
          if (w.videoUrl.startsWith("/api/videos/")) {
             await deleteVideo(w.videoUrl);
             await fetch(`/api/works/${encodeURIComponent(w.id)}`, { method: "DELETE" });
          } else if (!w.videoUrl.startsWith("http")) {
             const { error } = await supabase.functions.invoke("delete-video", {
               body: { path: w.videoUrl }
             });
             if (error) throw error;
          }
          setWorks(prev=>prev.filter(item=>item.id!==w.id));
          setEvals(prev=>prev.filter(item=>item.workId!==w.id));
          setLikes(prev=>prev.filter(item=>item.workId!==w.id));
          setShowCodePrompt(false); setCodeInput(""); setCodeErr("");
        } catch (err: any) {
          setCodeErr("❌ Delete Error: " + err.message);
        }
      }
    }else{
      setCodeErr("Invalid evaluation code • كود التصحيح غير صحيح");
    }
  }

  const critList=comp?.criteria?.length?comp.criteria:DEFAULT_CRITERIA;
  const teacher=users.find(u=>u.id===stu?.teacherId&&u.role==="teacher");
  const isWinner=certs.some(c=>c.workId===w.id);
  const[starPrompt,setStarPrompt]=useState(false);
  const[starCode,setStarCode]=useState("");
  const[starErr,setStarErr]=useState("");
  const[starOk,setStarOk]=useState(false);

  function confirmStar(){
    if(!comp)return;
    const ok=onPromote(w.id,w.competitionId,starCode);
    if(ok){setStarOk(true);setStarErr("");setStarPrompt(false);setStarCode("")}
    else{setStarErr("Invalid evaluation code • كود خاطئ")}
  }

  return(
    <article className="group rounded-[2rem] border border-white/10 bg-white/5 backdrop-blur-sm overflow-hidden flex flex-col hover:-translate-y-0.5 transition-all">
      <div className="relative bg-black">
        <WorkVideo w={w} users={users} setWorks={setWorks}/>
        {(isWinner||starOk)&&<div className="absolute top-2 right-2 bg-[#FDB51F] text-slate-900 text-xs font-black px-3 py-1 rounded-full shadow-lg">🏆 Winner</div>}
      </div>
      <div className="p-5 flex-1 flex flex-col">
        {comp&&<div className="text-[10px] font-black uppercase tracking-[0.22em] text-[#1697E8] mb-2">{comp.nameEn}</div>}
        <h3 className="text-lg font-black text-white leading-snug">{w.title}</h3>
        {w.desc&&<p className="mt-2 text-sm text-white/50 leading-relaxed line-clamp-2">{w.desc}</p>}
        <div className="mt-4 pt-4 border-t border-white/10 flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center rounded-full bg-[#1F2E76] text-sm font-black text-[#FDB51F] shrink-0">{stu.fullName[0]}</div>
          <div>
            <p className="text-sm font-black text-white">{stu.fullName}</p>
            <p className="text-xs text-white/40">{stu.grade}{teacher?` • Teacher: ${teacher.fullName}`:""}</p>
          </div>
          {/* Gold Star promote button */}
          {!isApproved(w)?(
            <span className="ml-auto text-[10px] font-bold text-white/30">⭐ after approval</span>
          ):!isWinner&&!starOk?(
            <button onClick={()=>setStarPrompt(true)} title="Promote to Winners • تحويل للفائزين" className="ml-auto text-2xl hover:scale-125 transition-transform opacity-50 hover:opacity-100">⭐</button>
          ):(
            <div className="ml-auto text-2xl" title="Winner!">🌟</div>
          )}
        </div>

        {/* Star Prompt Modal */}
        {starPrompt&&(
          <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 p-4 backdrop-blur-sm" onClick={()=>setStarPrompt(false)}>
            <div className="w-full max-w-sm rounded-3xl border border-white/10 bg-[#0f1535] p-6 shadow-2xl" onClick={e=>e.stopPropagation()}>
              <div className="text-center mb-5"><div className="text-3xl mb-2">⭐</div><h3 className="text-lg font-black text-white">Promote to Winners Board</h3><p className="text-sm text-white/50 mt-1">تحويل الطالب للوحة الفائزين وإصدار الشهادة</p><p className="text-sm font-bold text-[#FDB51F] mt-1">{stu.fullName}</p></div>
              <div className="space-y-4">
                <input value={starCode} onChange={e=>{setStarCode(e.target.value);setStarErr("")}} onKeyDown={e=>e.key==="Enter"&&confirmStar()} placeholder="Evaluation Code" className="w-full rounded-2xl border border-white/10 bg-slate-950/35 px-4 py-3 text-white text-center text-lg font-black tracking-[0.3em] uppercase outline-none placeholder:text-white/20 placeholder:tracking-normal placeholder:font-normal focus:border-[#FDB51F]" autoFocus/>
                {starErr&&<div className="rounded-xl bg-rose-500/10 border border-rose-400/20 px-4 py-3 text-sm text-rose-200 font-bold text-center">{starErr}</div>}
                <div className="flex gap-3">
                  <button onClick={()=>{setStarPrompt(false);setStarCode("");setStarErr("")}} className="flex-1 rounded-2xl border border-white/10 bg-white/5 py-3 text-sm font-bold text-white/60">Cancel</button>
                  <button onClick={confirmStar} className="flex-1 rounded-2xl bg-[#FDB51F] py-3 text-sm font-black text-slate-950 hover:bg-[#F7C84A]">🌟 Promote & Issue Cert</button>
                </div>
              </div>
            </div>
          </div>
        )}
        <div className="mt-auto pt-4 flex flex-col gap-2">
          {!isApproved(w)?(
            <div className="w-full rounded-2xl bg-white/5 border border-white/10 px-4 py-3 text-sm font-bold text-[#FFE29A]/80 text-center">⏳ Awaiting teacher approval before evaluation • بانتظار الاعتماد</div>
          ):ev?(
            <button onClick={()=>setViewEval(true)} className="w-full rounded-2xl bg-emerald-500/15 border border-emerald-400/20 px-4 py-3 text-sm font-black text-emerald-200 hover:bg-emerald-500/25 transition">
              📊 View Score ({tot(ev.scores)}/100) • مشاهدة التقييم
            </button>
          ):(
            <button onClick={()=>{setCodeMode("eval");setShowCodePrompt(true)}} className="w-full rounded-2xl bg-white/5 border border-white/10 px-4 py-3 text-sm font-bold text-white/40 hover:text-[#FDB51F] hover:border-[#FDB51F]/30 hover:bg-[#FDB51F]/10 transition cursor-pointer text-center">
              ⏳ Under Review • بانتظار التقييم — Click to Evaluate
            </button>
          )}
          {/* Edit existing eval */}
          {isApproved(w)&&ev&&(
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
            </div>
            <div className="space-y-4">
              <input value={codeInput} onChange={e=>{setCodeInput(e.target.value);if(codeErr)setCodeErr("")}} onKeyDown={e=>e.key==="Enter"&&verifyCode()} placeholder="Evaluation Code" className="w-full rounded-2xl border border-white/10 bg-slate-950/35 px-4 py-3 text-white text-center text-lg font-black tracking-[0.3em] uppercase outline-none placeholder:text-white/20 focus:border-[#FDB51F]" autoFocus/>
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
   ADMIN-ONLY COMPETITION TOOLS
   ════════════════════════════════════ */
function AdminUnlockModal({adminPw,onUnlock,onClose}:{adminPw:string;onUnlock:()=>void;onClose:()=>void}){
  const[pw,setPw]=useState(""); const[err,setErr]=useState("");
  function sub(e:React.FormEvent){
    e.preventDefault();
    if(pw.trim()!==adminPw){setErr("كلمة المرور غير صحيحة • Incorrect admin password");return}
    onUnlock();
  }
  return(
    <div className="fixed inset-0 z-[70] flex items-center justify-center bg-black/85 p-4 backdrop-blur-md" onClick={onClose}>
      <div className="w-full max-w-sm overflow-hidden rounded-[2rem] border border-white/10 bg-[#0f1535] shadow-2xl" onClick={e=>e.stopPropagation()}>
        <div className="h-1.5 w-full bg-gradient-to-r from-[#FDB51F] via-[#1697E8] to-[#10B09B]"/>
        <div className="p-6">
          <div className="mb-5 text-center">
            <div className="mx-auto mb-3 grid h-14 w-14 place-items-center rounded-2xl bg-[#FDB51F]/15 text-2xl ring-1 ring-[#FDB51F]/30">🔐</div>
            <h3 className="text-lg font-black text-white">Admin Access</h3>
            <p className="mt-1 text-sm text-white/50">دخول الإدارة — للحذف والتعديل</p>
          </div>
          <form onSubmit={sub} className="space-y-3">
            <Inp type="password" value={pw} onChange={e=>{setPw(e.target.value);setErr("")}} placeholder="Admin password" autoFocus/>
            {err&&<div className="rounded-xl border border-rose-400/20 bg-rose-500/10 px-4 py-3 text-sm font-bold text-rose-200">{err}</div>}
            <div className="flex gap-3">
              <button type="button" onClick={onClose} className="flex-1 rounded-2xl border border-white/10 bg-white/5 py-3 text-sm font-black text-white/70 hover:bg-white/10">Cancel</button>
              <button type="submit" className="flex-1 rounded-2xl bg-[#FDB51F] py-3 text-sm font-black text-slate-950 hover:bg-[#F7C84A]">Unlock 🔓</button>
            </div>
          </form>
        </div>
      </div>
    </div>
  );
}

function CompEditModal({comp,onClose,onSave}:{comp:Comp;onClose:()=>void;onSave:(patch:Partial<Comp>)=>void}){
  const[nEn,setNEn]=useState(comp.nameEn);
  const[nAr,setNAr]=useState(comp.nameAr);
  const[cClose,setCClose]=useState(comp.closingDate||"");
  const[cCert,setCCert]=useState(comp.certDate||"");
  const[err,setErr]=useState("");
  function save(e:React.FormEvent){
    e.preventDefault();
    if(!nEn.trim()||!nAr.trim()){setErr("أدخل الاسم باللغتين • Enter both names");return}
    if(cClose&&cCert&&cCert<cClose){setErr("تاريخ الشهادات يجب أن يكون بعد الإغلاق • Certificates date must follow the closing date");return}
    onSave({nameEn:nEn.trim(),nameAr:nAr.trim(),closingDate:cClose,certDate:cCert});
    onClose();
  }
  return(
    <div className="fixed inset-0 z-[70] flex items-center justify-center bg-black/85 p-4 backdrop-blur-md" onClick={onClose}>
      <div className="max-h-[90vh] w-full max-w-lg overflow-y-auto rounded-[2rem] border border-white/10 bg-[#0f1535] shadow-2xl" onClick={e=>e.stopPropagation()}>
        <div className="h-1.5 w-full bg-gradient-to-r from-[#1697E8] via-[#FDB51F] to-[#10B09B]"/>
        <div className="p-6">
          <div className="mb-5 flex items-start justify-between gap-4">
            <div><h3 className="text-xl font-black text-white">✏️ Edit Competition</h3><p className="mt-0.5 text-sm text-white/45">تعديل المسابقة — Admin only</p></div>
            <button onClick={onClose} className="rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm font-bold text-white hover:bg-white/10">✕</button>
          </div>
          <form onSubmit={save} className="space-y-4">
            <div><FL>Name (English)</FL><Inp value={nEn} onChange={e=>setNEn(e.target.value)}/></div>
            <div><FL>Name (Arabic) • الاسم بالعربية</FL><Inp value={nAr} onChange={e=>setNAr(e.target.value)} dir="rtl"/></div>
            <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
              <div><FL>Closing Date • تاريخ الإغلاق</FL><Inp type="date" value={cClose} onChange={e=>setCClose(e.target.value)}/></div>
              <div><FL>Certs Date • تاريخ الشهادات</FL><Inp type="date" value={cCert} onChange={e=>setCCert(e.target.value)}/></div>
            </div>
            {err&&<div className="rounded-xl border border-rose-400/20 bg-rose-500/10 px-4 py-3 text-sm font-bold text-rose-200">{err}</div>}
            <div className="flex gap-3 pt-1">
              <button type="button" onClick={onClose} className="flex-1 rounded-2xl border border-white/10 bg-white/5 py-3.5 text-sm font-black text-white/70 hover:bg-white/10">Cancel</button>
              <button type="submit" className="flex-[2] rounded-2xl bg-[#FDB51F] py-3.5 text-sm font-black text-slate-950 hover:bg-[#F7C84A]">Save Changes ✓</button>
            </div>
          </form>
        </div>
      </div>
    </div>
  );
}

function CompDeleteModal({comp,adminPw,worksCount,onClose,onConfirm}:{comp:Comp;adminPw:string;worksCount:number;onClose:()=>void;onConfirm:()=>void}){
  const[pw,setPw]=useState(""); const[err,setErr]=useState("");
  function confirm(e:React.FormEvent){
    e.preventDefault();
    if(pw.trim()!==adminPw){setErr("كلمة مرور الإدارة غير صحيحة • Incorrect admin password");return}
    onConfirm();
  }
  return(
    <div className="fixed inset-0 z-[70] flex items-center justify-center bg-black/85 p-4 backdrop-blur-md" onClick={onClose}>
      <div className="w-full max-w-md overflow-hidden rounded-[2rem] border border-rose-400/25 bg-[#0f1535] shadow-2xl" onClick={e=>e.stopPropagation()}>
        <div className="h-1.5 w-full bg-gradient-to-r from-rose-600 via-[#E21455] to-[#FDB51F]"/>
        <div className="p-6">
          <div className="mb-4 text-center">
            <div className="mx-auto mb-3 grid h-14 w-14 place-items-center rounded-2xl bg-rose-500/15 text-2xl ring-1 ring-rose-400/30">🗑️</div>
            <h3 className="text-xl font-black text-white">Delete Competition</h3>
            <p className="mt-1 text-sm text-white/50">حذف المسابقة نهائياً</p>
          </div>
          <div className="mb-4 rounded-2xl border border-white/10 bg-white/5 p-4">
            <p className="font-black text-white">{comp.nameEn}</p>
            <p className="text-sm font-bold text-[#FDB51F]" dir="rtl">{comp.nameAr}</p>
            <div className="mt-2 flex flex-wrap gap-2"><Pill t="pink">{worksCount} submissions</Pill><Pill t="slate">Closes: {comp.closingDate||"—"}</Pill></div>
          </div>
          <div className="mb-4 rounded-2xl border border-rose-400/20 bg-rose-500/10 p-4 text-xs font-bold leading-6 text-rose-200">
            ⚠️ تُحذف المسابقة نهائياً مع كل مشاركاتها وتقييماتها وشهاداتها، وتُزال من لوحة الظهور والأرشيف نهائياً.
            <br/>This permanently deletes the competition with all its submissions, evaluations and certificates. It is removed from the showcase board and the archive — nothing is kept.
          </div>
          <form onSubmit={confirm} className="space-y-3">
            <div><FL>Admin Password • كلمة مرور الإدارة</FL><Inp type="password" value={pw} onChange={e=>{setPw(e.target.value);setErr("")}} autoFocus/></div>
            {err&&<div className="rounded-xl border border-rose-400/25 bg-rose-500/15 px-4 py-3 text-sm font-bold text-rose-100">{err}</div>}
            <div className="flex gap-3">
              <button type="button" onClick={onClose} className="flex-1 rounded-2xl border border-white/10 bg-white/5 py-3.5 text-sm font-black text-white/70 hover:bg-white/10">Cancel</button>
              <button type="submit" className="flex-1 rounded-2xl bg-rose-600 py-3.5 text-sm font-black text-white hover:bg-rose-500">Delete Forever 🗑️</button>
            </div>
          </form>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════
   TEACHER LOGIN PAGE
   ════════════════════════════════════ */
function TeacherLoginPage({users,onBack,onOk}:{users:User[];onBack:()=>void;onOk:(t:User)=>void}){
  const[u,setU]=useState(""); const[p,setP]=useState(""); const[err,setErr]=useState(""); const[busy,setBusy]=useState(false);
  const norm=(x:any)=>String(x||"").trim().toLowerCase();
  const mapRow=(r:any):User=>({id:r.id,fullName:r.full_name||r.username||"Teacher",role:"teacher",grade:r.grade||"",section:r.section||"",teacherId:undefined,username:r.username,password:String(r.password??""),createdAt:r.created_at||"",active:r.active!==false});
  async function sub(e:React.FormEvent){
    e.preventDefault();
    setErr("");
    const username=u.trim(); const password=p;
    if(!username||!password){setErr("أدخل اسم المستخدم وكلمة المرور • Enter username and password");return}
    setBusy(true);
    try{
      let t=users.find(x=>x.role==="teacher"&&norm(x.username)===norm(username)&&String(x.password)===password&&x.active!==false);
      if(!t && isSupabaseConfigured){
        const{data,error}=await supabase.from("accounts").select("*").ilike("username",username).limit(20);
        if(error) console.warn("teacher lookup:",error.message);
        const row=((data||[]) as any[]).find(r=>String(r.password??"")===password && r.role==="teacher" && r.active!==false);
        if(row) t=mapRow(row);
        else if(((data||[]) as any[]).some(r=>norm(r.username)===norm(username))){setErr("كلمة المرور غير صحيحة • Incorrect password");setBusy(false);return}
      }
      if(!t){setErr("اسم المستخدم غير موجود أو ليس حساب معلم • Username not found or not a teacher account");setBusy(false);return}
      onOk(t);
    }catch(e:any){setErr("تعذر تسجيل الدخول • Login error: "+(e?.message||"unknown"))}
    finally{setBusy(false)}
  }
  return(
    <div className="relative flex min-h-screen items-center justify-center overflow-hidden bg-[linear-gradient(160deg,#071126,#0b1b3a,#0d1230)] px-5 py-10 text-white"><Bg/>
      <div className="pointer-events-none absolute -top-24 left-1/2 h-72 w-72 -translate-x-1/2 rounded-full bg-[#1697E8]/25 blur-3xl"/>
      <div className="relative w-full max-w-md overflow-hidden rounded-[2rem] border border-white/10 bg-white/[0.04] shadow-[0_30px_90px_-30px_rgba(22,151,232,0.6)] backdrop-blur-xl">
        <div className="h-1.5 w-full bg-gradient-to-r from-[#1697E8] via-[#10B09B] to-[#FDB51F]"/>
        <div className="p-7 sm:p-8">
          <button onClick={onBack} className="mb-5 text-xs font-bold text-white/45 transition hover:text-white">← Back • رجوع</button>
          <div className="mb-6 flex flex-col items-center text-center">
            <Logo s/>
            <div className="mt-4 inline-flex items-center gap-2 rounded-full border border-[#1697E8]/30 bg-[#1697E8]/15 px-3 py-1 text-[10px] font-black uppercase tracking-[0.2em] text-sky-200">👩‍🏫 Teacher Portal</div>
            <h1 className="mt-3 text-2xl font-black text-white">Teacher Login</h1>
            <p className="mt-1 text-sm font-bold text-[#FDB51F]" dir="rtl">دخول المعلم</p>
            <p className="mt-2 max-w-xs text-[11px] leading-5 text-white/40">Review and approve student submissions, then upload your own — same cloud account on every device.</p>
          </div>
          <form onSubmit={sub} className="space-y-4">
            <div><FL>Username • اسم المستخدم</FL><Inp value={u} onChange={e=>setU(e.target.value)} placeholder="e.g. heba" autoComplete="username"/></div>
            <div><FL>Password • كلمة المرور</FL><Inp type="password" value={p} onChange={e=>setP(e.target.value)} autoComplete="current-password"/></div>
            {err&&<div className="rounded-2xl border border-rose-400/20 bg-rose-500/10 px-4 py-3 text-sm font-bold leading-6 text-rose-200">{err}</div>}
            <button type="submit" disabled={busy} className={cn("w-full rounded-2xl py-4 text-base font-black shadow-lg shadow-[#1697E8]/20 transition",busy?"cursor-not-allowed bg-white/10 text-white/30":"bg-gradient-to-r from-[#1697E8] to-[#10B09B] text-white hover:brightness-110")}>{busy?"Signing in…":"Login 👩‍🏫"}</button>
          </form>
          <p className="mt-5 text-center text-[11px] leading-5 text-white/35">هل أنت طالب؟ <button onClick={onBack} className="font-black text-[#FDB51F] underline decoration-dotted">دخول الطالب</button></p>
        </div>
      </div>
    </div>
  );
}

/* ════════════════════════════════════
   TEACHER PORTAL (approvals + upload)
   ════════════════════════════════════ */
function TeacherPortalPage({teacher,users,works,comps,setWorks,onUpload,onBack}:{teacher:User;users:User[];works:Work[];comps:Comp[];setWorks:React.Dispatch<React.SetStateAction<Work[]>>;onUpload:()=>void;onBack:()=>void}){
  const[tab,setTab]=useState<"pending"|"approved"|"all">("pending");
  const[selClass,setSelClass]=useState<{grade:string;section:string}|null>(null);
  const[showAllClasses,setShowAllClasses]=useState(false);

  /* every student account assigned to this teacher */
  const myStudents=useMemo(()=>users.filter(u=>u.role==="student"&&u.teacherId===teacher.id),[users,teacher.id]);

  /* class groups = grade + section, each with its own submission counters */
  const classes=useMemo(()=>{
    const map=new Map<string,{grade:string;section:string;studentIds:string[];works:number;pending:number;approved:number}>();
    /* classes assigned to the teacher account itself (may be several) */
    decodeClasses(teacher).forEach(c=>{
      const key=`${c.grade}||${c.section||"—"}`;
      if(!map.has(key)) map.set(key,{grade:c.grade,section:c.section||"—",studentIds:[],works:0,pending:0,approved:0});
    });
    myStudents.forEach(st=>{
      const grade=st.grade||"Grade —";
      const section=st.section||"—";
      const key=`${grade}||${section}`;
      if(!map.has(key)) map.set(key,{grade,section,studentIds:[],works:0,pending:0,approved:0});
      map.get(key)!.studentIds.push(st.id);
    });
    map.forEach(g=>{
      works.forEach(w=>{
        if(!g.studentIds.includes(w.studentId)) return;
        g.works+=1;
        if(isApproved(w)) g.approved+=1; else g.pending+=1;
      });
    });
    const order=(g:string)=>{const n=parseInt(String(g).replace(/[^0-9]/g,""),10);return isNaN(n)?99:n;};
    return Array.from(map.values()).sort((a,b)=>order(a.grade)-order(b.grade)||String(a.section).localeCompare(String(b.section)));
  },[myStudents,works,teacher]);

  const classIds=selClass?myStudents.filter(st=>(st.grade||"Grade —")===selClass.grade&&(st.section||"—")===selClass.section).map(st=>st.id):null;
  const scoped = classIds ? works.filter(w=>classIds.includes(w.studentId)) : works;
  const pending=useMemo(()=>scoped.filter(w=>!isApproved(w)).sort((a,b)=>b.at.localeCompare(a.at)),[scoped]);
  const approved=useMemo(()=>scoped.filter(w=>isApproved(w)).sort((a,b)=>b.at.localeCompare(a.at)),[scoped]);
  const list = tab==="pending"?pending:tab==="approved"?approved:scoped;

  function approve(w:Work){
    if(w.studentId===teacher.id) return;
    setWorks(p=>p.map(x=>x.id===w.id?{...x,status:"approved" as WorkStatus,approvedBy:teacher.fullName,approvedAt:new Date().toISOString()}:x));
  }

  const gradeIcon=(g:string)=>{
    const n=parseInt(String(g).replace(/[^0-9]/g,""),10);
    if(!isNaN(n)) return n<=2?"🧸":n<=5?"📚":n<=9?"🔬":"🎓";
    return "📚";
  };

  return(
    <div className="relative min-h-screen bg-[linear-gradient(180deg,#0b1028,#0f1a3d,#0d1230)] text-white"><Bg/>
      <header className="sticky top-0 z-30 border-b border-white/[0.07] bg-[#0b1028]/85 backdrop-blur-xl">
        <div className="mx-auto flex max-w-6xl flex-col gap-3 px-5 py-4 sm:flex-row sm:items-center sm:justify-between sm:px-6">
          <div className="flex min-w-0 items-center gap-3">
            <Logo s/>
            <div className="min-w-0">
              <div className="text-[10px] font-black uppercase tracking-[0.28em] text-sky-300">Teacher Portal</div>
              <h1 className="truncate text-base font-black leading-tight sm:text-lg">{teacher.fullName}</h1>
              <p className="truncate text-[11px] font-semibold text-white/40">@{teacher.username} • {classes.length} classes • {myStudents.length} students</p>
              {classes.length>0&&<p className="truncate text-[10px] font-bold text-sky-300/70">{classes.map(c=>classLabel(c)).join(" · ")}</p>}
            </div>
          </div>
          <div className="flex gap-2">
            <button onClick={onUpload} className="flex-1 rounded-full bg-[#FDB51F] px-4 py-2.5 text-xs font-black text-slate-950 shadow-lg shadow-[#FDB51F]/20 transition hover:bg-[#F7C84A] sm:flex-none sm:text-sm">📹 Upload Video</button>
            <button onClick={onBack} className="flex-1 rounded-full border border-white/10 bg-white/5 px-4 py-2.5 text-xs font-bold text-white transition hover:bg-white/10 sm:flex-none sm:text-sm">← Logout</button>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-6xl px-5 py-7 sm:px-6">
        {!selClass ? (
          <>
            {/* ── overview ── */}
            <div className="mb-6 flex flex-wrap items-center justify-between gap-3">
              <div>
                <h2 className="text-xl font-black text-white sm:text-2xl">My Classes • صفوفي</h2>
                <p className="mt-0.5 text-xs font-semibold text-white/40">اختر صفاً لعرض أعمال طلابه المعتمد لها</p>
              </div>
              <div className="flex gap-2">
                <Pill t="gold">⏳ {works.filter(w=>myStudents.some(s=>s.id===w.studentId)&&!isApproved(w)).length} pending</Pill>
                <Pill t="green">✅ {works.filter(w=>myStudents.some(s=>s.id===w.studentId)&&isApproved(w)).length} approved</Pill>
              </div>
            </div>

            <div className="mb-8 grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-4">
              {classes.map(c=>{
                const co=getClassColor(c.grade);
                return(
                  <button key={`${c.grade}-${c.section}`} onClick={()=>{setSelClass({grade:c.grade,section:c.section});setTab("pending");}}
                    className="group relative overflow-hidden rounded-3xl border border-white/[0.08] p-5 text-left transition hover:-translate-y-0.5 hover:border-white/20"
                    style={{background:`linear-gradient(150deg,${co.bg},rgba(15,23,42,0.75))`}}>
                    <span className="absolute -right-6 -top-6 h-20 w-20 rounded-full opacity-30 blur-2xl" style={{background:co.border}} aria-hidden="true"/>
                    <div className="relative">
                      <div className="grid h-12 w-12 place-items-center rounded-2xl bg-black/25 text-2xl ring-1 ring-white/15">{gradeIcon(c.grade)}</div>
                      <h3 className="mt-3 text-base font-black leading-tight text-white">{c.grade}</h3>
                      <p className="text-xs font-bold text-white/55">Section {c.section}</p>
                      <div className="mt-3 flex flex-wrap gap-1.5 text-[10px] font-black">
                        <span className="rounded-full bg-black/30 px-2 py-1 text-white/70">👥 {c.studentIds.length}</span>
                        <span className="rounded-full bg-black/30 px-2 py-1 text-sky-200">🎬 {c.works}</span>
                        {c.pending>0&&<span className="rounded-full bg-[#FDB51F]/25 px-2 py-1 text-[#FFE29A]">⏳ {c.pending}</span>}
                      </div>
                      <span className="mt-3 block text-[11px] font-black text-white/40 transition group-hover:text-white">Open class →</span>
                    </div>
                  </button>
                );
              })}
              {classes.length===0&&(
                <div className="col-span-2 rounded-3xl border border-white/[0.08] bg-white/[0.03] p-8 text-center sm:col-span-3 lg:col-span-4">
                  <p className="text-3xl mb-2">🏫</p>
                  <p className="text-sm font-bold text-white/50">لا توجد صفوف مسندة إلى حسابك بعد</p>
                  <p className="mt-1 text-[11px] text-white/35">No classes are assigned to your account yet — ask the administrator.</p>
                  <button onClick={()=>setShowAllClasses(true)} className="mt-4 rounded-full border border-white/15 bg-white/5 px-4 py-2 text-xs font-black text-white/70 hover:bg-white/10">📂 Browse all submissions</button>
                </div>
              )}
            </div>

            {classes.length>0&&(
              <button onClick={()=>setShowAllClasses(v=>!v)} className="mb-4 rounded-full border border-white/10 bg-white/5 px-4 py-2 text-xs font-black text-white/60 hover:bg-white/10">{showAllClasses?"▲ Hide all submissions":"📂 All submissions • كل المشاركات"}</button>
            )}

            {showAllClasses&&(
              <div className="grid gap-5 md:grid-cols-2 xl:grid-cols-3">
                {works.slice().sort((a,b)=>b.at.localeCompare(a.at)).map(w=>{
                  const stu=users.find(x=>x.id===w.studentId);
                  const comp=comps.find(c=>c.id===w.competitionId);
                  const own=w.studentId===teacher.id;
                  return(
                    <article key={w.id} className="overflow-hidden rounded-2xl border border-white/[0.07] bg-white/[0.03]">
                      <SecureVideo path={w.videoUrl} poster={w.thumb} controls className="w-full aspect-video object-cover bg-black"/>
                      <div className="p-4">
                        <div className="mb-2 flex flex-wrap items-center gap-2"><StatusPill w={w}/>{own&&<span className="rounded-full border border-sky-400/25 bg-sky-500/15 px-2 py-0.5 text-[10px] font-black text-sky-200">Yours</span>}</div>
                        <div className="text-[10px] font-black uppercase tracking-[0.18em] text-[#1697E8]">{comp?.nameEn||"—"}</div>
                        <h3 className="mt-1 text-sm font-black leading-snug text-white">{w.title}</h3>
                        <p className="mt-1 text-xs font-bold text-[#FDB51F]">👤 {stu?.fullName||"—"}</p>
                        <p className="text-[11px] text-white/35">{stu?.grade||""}{stu?.section?` • ${stu.section}`:""}</p>
                        {!isApproved(w)?(
                          <button disabled={own} onClick={()=>approve(w)} className={cn("mt-3 w-full rounded-2xl py-3 text-xs font-black transition",own?"cursor-not-allowed bg-white/5 text-white/25":"bg-emerald-500 text-slate-950 hover:bg-emerald-400")}>{own?"لا يمكنك اعتماد فيديوك":"✅ Approve • اعتماد"}</button>
                        ):(
                          <div className="mt-3 rounded-2xl border border-emerald-400/20 bg-emerald-500/10 px-3 py-2.5 text-[11px] font-bold text-emerald-200">✅ Approved {w.approvedBy?`by ${w.approvedBy}`:""}</div>
                        )}
                      </div>
                    </article>
                  );
                })}
              </div>
            )}
          </>
        ) : (
          <>
            {/* ── class detail ── */}
            <div className="mb-5 flex flex-wrap items-center gap-3">
              <button onClick={()=>{setSelClass(null);setTab("pending");}} className="rounded-full border border-white/10 bg-white/5 px-4 py-2.5 text-xs font-black text-white hover:bg-white/10">← All classes</button>
              <div className="flex min-w-0 items-center gap-3">
                <span className="grid h-11 w-11 place-items-center rounded-2xl bg-white/10 text-xl ring-1 ring-white/15">{gradeIcon(selClass.grade)}</span>
                <div className="min-w-0">
                  <h2 className="truncate text-lg font-black leading-tight text-white">{selClass.grade} • Section {selClass.section}</h2>
                  <p className="text-[11px] font-semibold text-white/40">{(classIds||[]).length} students • {scoped.length} submissions</p>
                </div>
              </div>
            </div>

            <div className="mb-5 flex flex-wrap gap-2 rounded-2xl border border-white/[0.08] bg-white/[0.03] p-1.5">
              {([{k:"pending",t:`⏳ Pending (${pending.length})`},{k:"approved",t:`✅ Approved (${approved.length})`},{k:"all",t:`📁 All (${scoped.length})`}] as const).map(t=>(
                <button key={t.k} onClick={()=>setTab(t.k)} className={cn("rounded-xl px-4 py-2.5 text-xs font-black transition sm:text-sm",tab===t.k?"bg-[#FDB51F] text-slate-950":"text-white/50 hover:text-white")}>{t.t}</button>
              ))}
            </div>

            {list.length===0&&<div className="rounded-3xl border border-white/[0.08] bg-white/[0.03] p-14 text-center"><p className="mb-3 text-4xl">🎬</p><p className="text-sm font-bold text-white/40">لا توجد مشاركات في هذا الصف • Nothing here yet</p></div>}

            <div className="grid gap-5 md:grid-cols-2 xl:grid-cols-3">
              {list.map(w=>{
                const stu=users.find(x=>x.id===w.studentId);
                const comp=comps.find(c=>c.id===w.competitionId);
                const own=w.studentId===teacher.id;
                return(
                  <article key={w.id} className="overflow-hidden rounded-2xl border border-white/[0.07] bg-white/[0.03]">
                    <SecureVideo path={w.videoUrl} poster={w.thumb} controls className="w-full aspect-video object-cover bg-black"/>
                    <div className="p-4">
                      <div className="mb-2 flex flex-wrap items-center gap-2"><StatusPill w={w}/>{own&&<span className="rounded-full border border-sky-400/25 bg-sky-500/15 px-2 py-0.5 text-[10px] font-black text-sky-200">Yours</span>}</div>
                      <div className="text-[10px] font-black uppercase tracking-[0.18em] text-[#1697E8]">{comp?.nameEn||"—"}</div>
                      <h3 className="mt-1 text-sm font-black leading-snug text-white">{w.title}</h3>
                      <p className="mt-1 text-xs font-bold text-[#FDB51F]">👤 {stu?.fullName||"—"}</p>
                      <p className="text-[11px] text-white/35">{stu?.grade||""}{stu?.section?` • ${stu.section}`:""}</p>
                      {!isApproved(w)?(
                        <button disabled={own} onClick={()=>approve(w)} className={cn("mt-3 w-full rounded-2xl py-3 text-xs font-black transition",own?"cursor-not-allowed bg-white/5 text-white/25":"bg-emerald-500 text-slate-950 hover:bg-emerald-400")}>{own?"لا يمكنك اعتماد فيديوك":"✅ Approve • اعتماد"}</button>
                      ):(
                        <div className="mt-3 rounded-2xl border border-emerald-400/20 bg-emerald-500/10 px-3 py-2.5 text-[11px] font-bold text-emerald-200">✅ Approved {w.approvedBy?`by ${w.approvedBy}`:""}</div>
                      )}
                    </div>
                  </article>
                );
              })}
            </div>
          </>
        )}
      </main>
    </div>
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
  const[currentStudent,setCurrentStudent]=useState<User|null>(null);
  const[currentTeacher,setCurrentTeacher]=useState<User|null>(null);

  // Phase 2A foundation: initialize Supabase session handling without changing current UI behavior
  useEffect(() => {
    if (!isSupabaseConfigured) return;
    
    supabase.auth.getSession();
    const { data } = supabase.auth.onAuthStateChange(() => {
      // Session is handled by Supabase internally for future phases.
      // No UI behavior changes in Phase 2A.
    });
    return () => data.subscription.unsubscribe();
  }, []);

  // ═══ CLOUD SYNC: Load ALL data from Supabase on startup ═══
  useEffect(()=>{
    async function loadFromSupabase(){
      try {
        const {data:dbUsers} = await supabase.from("accounts").select("*");
        if(dbUsers && dbUsers.length > 0){
          const mapped = dbUsers.map((u:any)=>({
            id:u.id, fullName:u.full_name, role:u.role, grade:u.grade,
            section:u.section, teacherId:u.teacher_id, username:u.username,
            password:u.password, createdAt:u.created_at||"", active:u.active!==false
          }));
          setUsers(prev=>{
            const ids=new Set(mapped.map((u:any)=>u.id));
            const merged=[...mapped,...prev.filter(u=>!ids.has(u.id))];
            localStorage.setItem(SK.u,JSON.stringify(merged));
            return merged;
          });
        }
        const {data:dbWorks} = await supabase.from("works_cloud").select("*").order('at',{ascending:false});
        if(dbWorks && dbWorks.length > 0){
          const mapped = dbWorks.map((w:any)=>({
            id:w.id, competitionId:w.competition_id, studentId:w.student_id,
            title:w.title, desc:w.desc||"", videoUrl:w.video_url, thumb:w.thumb||"", at:w.at||new Date().toISOString(),
            status:(w.status==="approved"?"approved":"pending") as WorkStatus,
            approvedBy:w.approved_by||undefined, approvedAt:w.approved_at||undefined
          }));
          setWorks(prev=>{
            const ids=new Set(mapped.map((w:any)=>w.id));
            const merged=[...mapped,...prev.filter(w=>!ids.has(w.id))];
            localStorage.setItem(SK.w,JSON.stringify(merged));
            return merged;
          });
        }
        const {data:dbComps,error:compsErr} = await supabase.from("competitions_cloud").select("*");
        if(!compsErr && Array.isArray(dbComps)){
          const adminRow = dbComps.find((c:any)=>String(c.id)==="__admin__");
          if(adminRow?.name_en) setAdminPw(String(adminRow.name_en));
          const realRows = dbComps.filter((c:any)=>String(c.id)!=="__meta__" && !String(c.id).startsWith("__"));
          const initialized = dbComps.some((c:any)=>String(c.id)==="__meta__");
          if(realRows.length>0){
            const mapped = realRows.map((c:any)=>({
              id:c.id, nameEn:c.name_en||"Competition", nameAr:c.name_ar||"مسابقة",
              logoUrl:c.logo_url||"/gems-logo.png",
              grades:Array.isArray(c.grades)?c.grades:[],
              rulesAr:c.rules_ar||"", rulesEn:c.rules_en||"",
              closingDate:c.closing_date||"", certDate:c.cert_date||"",
              started:c.started!==false, archived:c.archived===true,
              criteria:Array.isArray(c.criteria)&&c.criteria.length?c.criteria:DEFAULT_CRITERIA,
              evalCode:c.eval_code||""
            })) as Comp[];
            setComps(mapped);
            try{ localStorage.setItem(SK.c,JSON.stringify(mapped)); }catch{}
          }else if(initialized){
            // the cloud was initialised and has no competitions left → every one was deleted
            setComps([]);
            try{ localStorage.setItem(SK.c,"[]"); }catch{}
          }
        }
        const {data:dbEvals} = await supabase.from("evaluations").select("*");
        if(dbEvals && dbEvals.length > 0){
          const mapped = dbEvals.map((e:any)=>({
            id:e.id, workId:e.work_id, scores:e.scores||[], at:e.at||new Date().toISOString(), approved:e.approved||false
          }));
          setEvals(prev=>{
            const ids=new Set(mapped.map((e:any)=>e.id));
            return [...mapped,...prev.filter(e=>!ids.has(e.id))];
          });
        }
        const {data:dbCerts} = await supabase.from("certificates").select("*");
        if(dbCerts && dbCerts.length > 0){
          const mapped = dbCerts.map((c:any)=>({
            id:c.id, workId:c.work_id, studentId:c.student_id, compId:c.comp_id,
            compName:c.comp_name, score:c.score, rank:c.rank, grade:c.grade,
            section:c.section, issuedAt:c.issued_at
          }));
          setCerts(prev=>{
            const ids=new Set(mapped.map((c:any)=>c.id));
            return [...mapped,...prev.filter(c=>!ids.has(c.id))];
          });
        }
      } catch(err){ console.error("Supabase load error:", err); }
    }
    loadFromSupabase();
  },[]);

  useEffect(()=>{sv(SK.u,users)},[users]);
  useEffect(()=>{sv(SK.pw,adminPw)},[adminPw]);

  /* cloud copy of the admin password so every device unlocks with the same one */
  useEffect(()=>{
    if(!isSupabaseConfigured || !adminPw) return;
    const id=setTimeout(async ()=>{
      try{
        await supabase.from("competitions_cloud").upsert({
          id:"__admin__", name_en:String(adminPw), name_ar:"__admin__", logo_url:"",
          grades:[], rules_ar:"", rules_en:"", closing_date:"", cert_date:"",
          started:true, archived:true, criteria:[], eval_code:""
        },{onConflict:"id"});
      }catch{}
    },400);
    return ()=>clearTimeout(id);
  },[adminPw]);
  useEffect(()=>{sv(SK.c,comps)},[comps]);
  useEffect(()=>{sv(SK.w,works)},[works]);
  useEffect(()=>{sv(SK.e,evals)},[evals]);
  useEffect(()=>{sv(SK.l,likes)},[likes]);
  useEffect(()=>{sv(SK.cert,certs)},[certs]);
  useEffect(()=>{sv(SK.arch,archive)},[archive]);

  function promoteToCertificate(workId:string, compId:string, evalCode:string):boolean{
    const comp=comps.find(c=>c.id===compId);
    if(!comp||evalCode.trim().toUpperCase()!==comp.evalCode?.toUpperCase())return false;
    const work=works.find(w=>w.id===workId); if(!work)return false;
    const stu=users.find(u=>u.id===work.studentId); if(!stu)return false;
    const ev=evals.find(e=>e.workId===workId);
    const score=ev?tot(ev.scores):0;
    const rank=certs.filter(c=>c.compId===compId).length+1;
    if(!certs.some(c=>c.workId===workId)){
      setCerts(p=>[...p,{id:mid("cert"),workId,studentId:stu.id,compId,compName:comp.nameEn,score,rank,grade:stu.grade,section:stu.section,issuedAt:new Date().toISOString().slice(0,10)}]);
    }
    return true;
  }

  function handleArchive(comp:Comp, winners:Array<{rank:number;studentName:string;grade:string;section:string;score:number}>){
    const entry:ArchiveEntry={id:mid("arch"),compId:comp.id,compName:comp.nameEn,compNameAr:comp.nameAr,archivedAt:new Date().toISOString().slice(0,10),winners};
    setArchive(p=>[entry,...p]);
    setComps(p=>p.map(c=>c.id===comp.id?{...c,archived:true}:c));
  }

  if(view==="home")return<HomePage onStart={()=>setView("showcase")} onAdmin={()=>setView("adminLogin")} onStudent={()=>setView("studentLogin")} onTeacher={()=>setView("teacherLogin")}/>;
  if(view==="showcase")return<ShowcasePage comps={comps} users={users} works={works} evals={evals} likes={likes} setLikes={setLikes} setEvals={setEvals} setWorks={setWorks} setComps={setComps} setCerts={setCerts} setArchive={setArchive} certs={certs} adminPw={adminPw} onPromote={promoteToCertificate} onBack={()=>setView("home")} onClass={c=>{setSelClass(c);setView("classBoard")}} onParticipate={id=>{setSelCompId(id);setView("studentLogin")}} onWinners={()=>setView("winners")} onAllWorks={()=>setView("allWorks")} />;
  if(view==="allWorks")return<GlobalWorksPage users={users} works={works} evals={evals} comps={comps} likes={likes} setLikes={setLikes} setEvals={setEvals} setWorks={setWorks} certs={certs} onPromote={promoteToCertificate} onBack={()=>setView("showcase")}/>;
  if(view==="winners")return<WinnersPage comps={comps} works={works} evals={evals} users={users} likes={likes} setLikes={setLikes} certs={certs} onPromote={promoteToCertificate} adminPw={adminPw} onArchive={handleArchive} onBack={()=>setView("showcase")}/>;
  if(view==="classBoard"&&selClass)return<ClassBoardPage cls={selClass} users={users} works={works} evals={evals} comps={comps} likes={likes} setLikes={setLikes} setEvals={setEvals} setWorks={setWorks} certs={certs} onPromote={promoteToCertificate} onBack={()=>setView("showcase")}/>;
  if(view==="adminLogin")return<AdminLoginPage adminPw={adminPw} onBack={()=>setView("home")} onOk={()=>setView("admin")}/>;
  if(view==="admin")return<AdminPage users={users} setUsers={setUsers} comps={comps} setComps={setComps} archive={archive} adminPw={adminPw} setAdminPw={setAdminPw} onBack={()=>setView("home")} />;
  if(view==="studentLogin")return<StudentLoginPage users={users} onBack={()=>setView("home")} onOk={s=>{ if(s.role==="teacher"){setCurrentTeacher(s);setCurrentStudent(s);setView("teacherPortal");} else {setCurrentStudent(s);setView("student");} }}/>;
  if(view==="teacherLogin")return<TeacherLoginPage users={users} onBack={()=>setView("home")} onOk={t=>{setCurrentTeacher(t);setCurrentStudent(t);setView("teacherPortal")}}/>;
  if(view==="teacherPortal"&&currentTeacher)return<TeacherPortalPage teacher={currentTeacher} users={users} works={works} comps={comps} setWorks={setWorks} onUpload={()=>{setCurrentStudent(currentTeacher);setView("student")}} onBack={()=>{setCurrentTeacher(null);setCurrentStudent(null);setView("home")}}/>;
  if(view==="teacherPortal"&&!currentTeacher)return<TeacherLoginPage users={users} onBack={()=>setView("home")} onOk={t=>{setCurrentTeacher(t);setCurrentStudent(t);setView("teacherPortal")}}/>;
  if(view==="student"&&currentStudent)return<StudentPage student={currentStudent} comps={comps} works={works} evals={evals} certs={certs} selCompId={selCompId} setWorks={setWorks} onBack={()=>{setCurrentStudent(null);setView("home")}}/>;
  if(view==="student"&&!currentStudent)return<StudentLoginPage users={users} onBack={()=>setView("home")} onOk={s=>{ if(s.role==="teacher"){setCurrentTeacher(s);setCurrentStudent(s);setView("teacherPortal");} else {setCurrentStudent(s);setView("student");} }}/>;
  return null;
}

/* ════════════════════════════════════
   1. HOME
   ════════════════════════════════════ */
function HomePage({onStart,onAdmin,onStudent,onTeacher}:{onStart:()=>void;onAdmin:()=>void;onStudent:()=>void;onTeacher:()=>void}){
  const pillars = [
    {ar:"اللغة العربية",en:"Arabic Language",icon:"📖",tone:"border-[#1697E8]/30 bg-[#1697E8]/15 text-sky-200"},
    {ar:"التربية الإسلامية",en:"Islamic Education",icon:"🕌",tone:"border-[#10B09B]/30 bg-[#10B09B]/15 text-teal-200"},
    {ar:"الدراسات الاجتماعية",en:"Social Studies",icon:"🌍",tone:"border-[#E21455]/30 bg-[#E21455]/15 text-pink-200"},
  ];
  const stats = [
    {num:"✦",en:"Competitions",ar:"المسابقات",c:"text-[#FDB51F]"},
    {num:"✦",en:"Class Galleries",ar:"لوحات الصفوف",c:"text-sky-300"},
    {num:"10",en:"Criteria",ar:"معايير التقييم",c:"text-emerald-300"},
  ];
  return(
  <div className="relative min-h-screen overflow-hidden bg-[linear-gradient(180deg,#0b1028,#0f1535,#0d1230)] text-white"><Bg/>

    {/* ── HEADER ── */}
    <header className="sticky top-0 z-30 border-b border-white/10 bg-[#0b1028]/85 backdrop-blur-xl">
      <div className="mx-auto flex max-w-7xl flex-col gap-4 px-5 py-4 sm:px-6 lg:flex-row lg:items-center lg:justify-between">
        <div className="flex items-center gap-3 sm:gap-4">
          <Logo s/>
          <div className="min-w-0">
            <p className="truncate text-sm font-black uppercase italic tracking-tight text-white sm:text-base">GWA Champions</p>
            <p className="truncate text-[11px] font-semibold text-[#FDB51F]/80 sm:text-xs" dir="rtl">أبطال اللغة العربية والتربية الإسلامية والدراسات الاجتماعية</p>
          </div>
        </div>
        <div className="flex items-center gap-2 sm:gap-3">
          <button onClick={onStudent} className="flex-1 rounded-full border border-white/15 bg-white/5 px-4 py-2.5 text-xs font-bold text-white backdrop-blur transition hover:bg-white/10 sm:flex-none sm:px-5 sm:text-sm">🎓 Student</button>
          <button onClick={onTeacher} className="flex-1 rounded-full border border-[#1697E8]/40 bg-[#1697E8]/15 px-4 py-2.5 text-xs font-black text-sky-100 transition hover:bg-[#1697E8]/25 sm:flex-none sm:px-5 sm:text-sm">👩‍🏫 Teacher</button>
          <button onClick={onAdmin} className="flex-1 rounded-full bg-[#FDB51F] px-4 py-2.5 text-xs font-black text-slate-950 shadow-lg shadow-[#FDB51F]/20 transition hover:bg-[#F7C84A] sm:flex-none sm:px-5 sm:text-sm">الإدارة • Admin</button>
        </div>
      </div>
    </header>

    <main className="relative z-10 mx-auto max-w-7xl px-5 pb-16 pt-8 sm:px-6 lg:pt-12">

      {/* ── HERO PHOTO: full, uncropped, taller banner ── */}
      <section aria-label="GWA Champions" className="relative mx-auto max-w-6xl">
        <div className="pointer-events-none absolute -inset-3 rounded-[2.5rem] bg-gradient-to-br from-[#1697E8]/25 via-[#FDB51F]/12 to-transparent blur-3xl sm:-inset-5" aria-hidden="true"/>
        <figure className="relative m-0">
          <div className="overflow-hidden rounded-[1.75rem] border border-[#FDB51F]/25 bg-[#0b1028] shadow-[0_30px_90px_-25px_rgba(22,151,232,0.5)] sm:rounded-[2.25rem]">
            <img
              src="/hero-student.png"
              alt="GWA Arabic and Islamic Champion holding the trophy"
              className="block h-auto w-full object-contain"
              loading="eager"
            />
          </div>
          <figcaption className="mt-3 flex flex-col items-center gap-4 rounded-[1.5rem] border border-white/10 bg-white/[0.04] px-5 py-4 text-center backdrop-blur-xl sm:mt-4 sm:flex-row sm:justify-between sm:rounded-[1.75rem] sm:px-7 sm:py-5 sm:text-left">
            <div className="flex items-center gap-3 sm:gap-4">
              <span className="grid h-12 w-12 shrink-0 place-items-center rounded-2xl bg-[#FDB51F]/15 text-2xl ring-1 ring-[#FDB51F]/30">🏆</span>
              <div>
                <p className="text-base font-black leading-tight tracking-tight text-white sm:text-2xl">Arabic &amp; Islamic Champions</p>
                <p className="mt-0.5 text-sm font-bold text-[#FFE29A] sm:text-lg" dir="rtl">أبطال العربية والإسلامية</p>
              </div>
            </div>
            <div className="flex flex-wrap items-center justify-center gap-2 sm:justify-end">
              <span className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-[10px] font-bold uppercase tracking-wider text-white/60">GEMS World Academy Dubai</span>
              <span className="rounded-full border border-white/10 bg-white/5 px-3 py-1.5 text-[10px] font-bold text-white/60" dir="rtl">دبي — الإمارات</span>
            </div>
          </figcaption>
        </figure>
      </section>

      {/* ── CONTENT ── */}
      <section className="mt-10 grid items-start gap-8 lg:mt-14 lg:grid-cols-[1.1fr_1fr] lg:gap-12">

        {/* copy */}
        <div>
          <div className="mb-5 inline-flex w-fit items-center gap-2 rounded-full border border-[#FDB51F]/25 bg-[#FDB51F]/10 px-4 py-1.5 sm:px-5 sm:py-2">
            <span className="h-2 w-2 animate-pulse rounded-full bg-[#FDB51F]"/>
            <span className="text-[10px] font-black uppercase tracking-[0.22em] text-[#FFE29A] sm:text-xs">GEMS World Academy Dubai</span>
          </div>
          <h2 className="text-3xl font-black leading-[1.15] tracking-tight text-white sm:text-4xl xl:text-5xl">
            Celebrating Our Students'
            <span className="mt-1 block bg-gradient-to-r from-[#FDB51F] via-[#1697E8] to-[#10B09B] bg-clip-text text-transparent">Outstanding Achievements</span>
          </h2>
          <p className="mt-4 max-w-xl text-sm leading-7 text-white/55 sm:text-base">A world-class bilingual platform showcasing student excellence in Arabic Language, Islamic Education &amp; Social Studies.</p>

          <div className="mt-5 rounded-2xl border border-white/10 bg-white/5 px-5 py-4 backdrop-blur-sm sm:px-6" dir="rtl">
            <p className="text-right text-sm font-bold leading-8 text-white/80 sm:text-base">أهلاً بكم في <span className="font-black text-[#FDB51F]">أكاديمية جيمس العالمية</span> — دبي، حيث نحتفل بإنجازات طلابنا الأبطال المتميزين في:</p>
            <div className="mt-3 flex flex-wrap justify-end gap-2">
              {pillars.map(p=>(
                <span key={p.en} title={p.en} className={cn("rounded-full border px-3 py-1.5 text-xs font-bold sm:text-sm",p.tone)}>{p.icon} {p.ar}</span>
              ))}
            </div>
          </div>
        </div>

        {/* actions + stats */}
        <div className="flex flex-col gap-6">
          <button onClick={onStart} className="group w-full rounded-full bg-gradient-to-r from-[#FDB51F] to-[#E21455] p-[2px] shadow-[0_0_40px_rgba(253,181,31,0.3)] transition-all hover:scale-[1.02] hover:shadow-[0_0_60px_rgba(253,181,31,0.45)]">
            <span className="flex items-center justify-center gap-3 rounded-full bg-[#0b1028] px-6 py-4">
              <span className="text-xl">🚀</span>
              <span className="text-base font-black text-white sm:text-lg">Start Viewing • بدء المشاهدة</span>
            </span>
          </button>

          <div className="grid grid-cols-3 gap-3">
            {stats.map(st=>(
              <div key={st.en} className="rounded-2xl border border-white/10 bg-white/5 p-4 text-center backdrop-blur-sm transition hover:border-white/20 hover:bg-white/[0.07]">
                <div className={cn("text-2xl font-black",st.c)}>{st.num}</div>
                <div className="mt-1 text-[10px] font-bold uppercase tracking-widest text-white/40">{st.en}</div>
                <div className="mt-1 text-xs font-bold text-white/50" dir="rtl">{st.ar}</div>
              </div>
            ))}
          </div>

          <div className="grid gap-3">
            {[
              {t:"Student Portal",ar:"بوابة الطالب",d:"ارفع فيديو مشاركتك وتابع حالة الاعتماد",icon:"🎓",ring:"border-[#FDB51F]/40 hover:bg-[#FDB51F]/10",chip:"bg-[#FDB51F]/15 text-[#FFE29A]",go:onStudent},
              {t:"Teacher Portal",ar:"بوابة المعلم",d:"راجع المشاركات واعتمدها ثم ارفع فيديوهاتك",icon:"👩‍🏫",ring:"border-[#1697E8]/40 hover:bg-[#1697E8]/10",chip:"bg-[#1697E8]/15 text-sky-200",go:onTeacher},
              {t:"Administration",ar:"لوحة الإدارة",d:"إنشاء المسابقات والتحكم الكامل بالمنصة",icon:"🛡️",ring:"border-white/15 hover:bg-white/10",chip:"bg-white/10 text-white/70",go:onAdmin},
            ].map(p=>(
              <button key={p.t} onClick={p.go} className={cn("group flex w-full items-center gap-3 rounded-2xl border bg-white/[0.03] px-4 py-3.5 text-left transition",p.ring)}>
                <span className={cn("grid h-11 w-11 shrink-0 place-items-center rounded-xl text-xl",p.chip)}>{p.icon}</span>
                <span className="min-w-0 flex-1">
                  <span className="block text-sm font-black leading-tight text-white">{p.t}</span>
                  <span className="block text-[11px] font-bold leading-5 text-white/40" dir="rtl">{p.ar} — {p.d}</span>
                </span>
                <span className="text-white/30 transition group-hover:translate-x-1 group-hover:text-white">→</span>
              </button>
            ))}
          </div>
        </div>
      </section>
    </main>
  </div>
  );
}

/* ════════════════════════════════════
   2. SHOWCASE
   ════════════════════════════════════ */
function ShowcasePage({comps,users,works,evals,likes,setLikes,setEvals,setWorks,setComps,setCerts,setArchive,certs,adminPw,onPromote,onBack,onParticipate,onWinners,onAllWorks}:{comps:Comp[];users:User[];works:Work[];evals:Ev[];likes:Like[];setLikes:React.Dispatch<React.SetStateAction<Like[]>>;setEvals:React.Dispatch<React.SetStateAction<Ev[]>>;setWorks:React.Dispatch<React.SetStateAction<Work[]>>;setComps:React.Dispatch<React.SetStateAction<Comp[]>>;setCerts:React.Dispatch<React.SetStateAction<Cert[]>>;setArchive:React.Dispatch<React.SetStateAction<ArchiveEntry[]>>;certs:Cert[];adminPw:string;onPromote:(wid:string,cid:string,code:string)=>boolean;onBack:()=>void;onClass:(c:ClassInfo)=>void;onParticipate:(id:string)=>void;onWinners:()=>void;onAllWorks:()=>void}){
  const evaluatedCount=useMemo(()=>works.filter(w=>isApproved(w)&&evals.some(e=>e.workId===w.id)).length,[works,evals]);
  /* ── ADMIN MODE (unlocked only with the admin password) ── */
  const[adminMode,setAdminMode]=useState(false);
  const[askPw,setAskPw]=useState(false);
  const[editComp,setEditComp]=useState<Comp|null>(null);
  const[delComp,setDelComp]=useState<Comp|null>(null);
  function deleteCompetitionForever(c:Comp){
    const doomed=works.filter(w=>w.competitionId===c.id).map(w=>w.id);
    setWorks(p=>p.filter(w=>w.competitionId!==c.id));
    setEvals(p=>p.filter(e=>!doomed.includes(e.workId)));
    setCerts(p=>p.filter(x=>x.compId!==c.id));
    setArchive(p=>p.filter(a=>a.compId!==c.id));
    setComps(p=>p.filter(x=>x.id!==c.id));
    setDelComp(null);
    try{
      if(isSupabaseConfigured){
        supabase.from("evaluations").delete().in("work_id",doomed).then(()=>{});
        supabase.from("works_cloud").delete().eq("competition_id",c.id).then(()=>{});
      }
    }catch{}
  }
  const active=useMemo(()=>comps.filter(c=>c.started&&!c.archived),[comps]);
  const[selCompId,setSelCompId]=useState<string|null>(null);
  const[gradeFilter,setGradeFilter]=useState<string|null>(null);
  const selComp=selCompId?comps.find(c=>c.id===selCompId):null;

  const compWorks=useMemo(()=>{
    if(!selComp)return[];
    let w=works.filter(x=>x.competitionId===selComp.id);
    if(gradeFilter){
      w=w.filter(x=>{const s=users.find(u=>u.id===x.studentId);return s?.grade===gradeFilter});
    }
    return w.sort((a,b)=>b.at.localeCompare(a.at));
  },[selComp,works,users,gradeFilter]);

  const compGrades=useMemo(()=>{
    if(!selComp)return[];
    const grades=new Set<string>();
    works.filter(w=>w.competitionId===selComp.id).forEach(w=>{
      const s=users.find(u=>u.id===w.studentId);
      if(s)grades.add(s.grade);
    });
    return Array.from(grades).sort();
  },[selComp,works,users]);

  if(!selComp){
    return(
      <div className="relative min-h-screen bg-[linear-gradient(180deg,#0d1230,#111a45,#0d1230)] text-white"><Bg/>
        <header className="sticky top-0 z-20 border-b border-white/10 bg-[#0f1535]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center">
          <div className="flex items-center gap-4"><Logo s/><div><h2 className="text-lg font-black uppercase italic">Active Competitions</h2><p className="text-sm text-white/40" dir="rtl">المسابقات النشطة</p></div></div>
          <div className="flex gap-3">
            <button onClick={()=>{ if(adminMode) setAdminMode(false); else setAskPw(true); }} className={cn("rounded-full border px-4 py-2.5 text-xs font-black transition",adminMode?"border-emerald-400/40 bg-emerald-500/20 text-emerald-200":"border-white/10 bg-white/5 text-white/70 hover:bg-white/10")}>{adminMode?"✅ Admin mode":"🔒 Admin"}</button>
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
                <div key={c.id} className="relative">
                <button onClick={()=>{setSelCompId(c.id);setGradeFilter(null)}} className="group w-full text-left rounded-[2rem] border border-white/10 p-6 hover:-translate-y-1 transition-all relative overflow-hidden backdrop-blur-sm" style={{background:`linear-gradient(140deg,${co.bg},rgba(15,23,42,0.7))`}}>
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
                {adminMode&&(
                  <div className="absolute right-3 top-3 z-20 flex gap-2">
                    <button onClick={()=>setEditComp(c)} className="rounded-full bg-sky-500 px-3 py-2 text-[11px] font-black text-white shadow-lg transition hover:bg-sky-400">✏️ Edit</button>
                    <button onClick={()=>setDelComp(c)} className="rounded-full bg-rose-600 px-3 py-2 text-[11px] font-black text-white shadow-lg transition hover:bg-rose-500">🗑 Delete</button>
                  </div>
                )}
                </div>
              );
            })}
            <button onClick={onAllWorks} className="group text-left rounded-[2rem] border border-white/10 bg-[linear-gradient(140deg,rgba(255,255,255,0.1),rgba(255,255,255,0.05))] p-8 hover:-translate-y-1 hover:border-white/30 transition-all relative overflow-hidden backdrop-blur-sm">
              <div className="absolute top-0 right-0 w-32 h-32 bg-white/5 blur-3xl rounded-full group-hover:bg-white/10 transition"/>
              <div className="relative z-10">
                <div className="text-xs font-black uppercase tracking-[0.24em] text-[#FDB51F] mb-2">Global Activity • كل المشاركات</div>
                <h3 className="text-3xl font-black text-white">All Submissions</h3>
                <div className="text-lg font-bold text-sky-400">View by upload date</div>
                <div className="mt-6 pt-4 border-t border-white/10 flex justify-between items-center text-sm font-bold text-white/50">
                  <span>🌎 All Grades</span>
                  <span className="text-white group-hover:scale-105 transition">View All →</span>
                </div>
              </div>
            </button>
          </div>
        </main>
        {askPw&&!adminMode&&<AdminUnlockModal adminPw={adminPw} onUnlock={()=>{setAdminMode(true);setAskPw(false);}} onClose={()=>setAskPw(false)}/>}
        {editComp&&<CompEditModal comp={editComp} onClose={()=>setEditComp(null)} onSave={p=>{const id=editComp.id;setComps(list=>list.map(c=>c.id===id?{...c,...p}:c));}}/>}
        {delComp&&<CompDeleteModal comp={delComp} adminPw={adminPw} worksCount={works.filter(w=>w.competitionId===delComp.id).length} onClose={()=>setDelComp(null)} onConfirm={()=>deleteCompetitionForever(delComp)}/>}
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
          {compWorks.map(w=><WorkCard key={w.id} w={w} users={users} evals={evals} comps={comps} likes={likes} setLikes={setLikes} setEvals={setEvals} setWorks={setWorks} certs={certs} onPromote={onPromote}/>)}
        </div>
      </main>
    </div>
  );
}

/* ════════════════════════════════════
   3. CLASS BOARD
   ════════════════════════════════════ */
function ClassBoardPage({cls,users,works,evals,comps,likes,setLikes,setEvals,setWorks,certs,onPromote,onBack}:{cls:ClassInfo;users:User[];works:Work[];evals:Ev[];comps:Comp[];likes:Like[];setLikes:React.Dispatch<React.SetStateAction<Like[]>>;setEvals:React.Dispatch<React.SetStateAction<Ev[]>>;setWorks:React.Dispatch<React.SetStateAction<Work[]>>;certs:Cert[];onPromote:(wid:string,cid:string,code:string)=>boolean;onBack:()=>void}){
  const classWorks=useMemo(()=>works.filter(w=>{const s=users.find(u=>u.id===w.studentId);return s&&s.grade===cls.grade&&s.section===cls.section;}),[cls,users,works]);
  return(
    <div className="relative min-h-screen bg-[#050916] text-white"><Bg/>
      <header className="sticky top-0 z-20 border-b border-white/10 bg-[#050916]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center">
        <button onClick={onBack} className="rounded-full border border-white/10 bg-white/5 px-5 py-2.5 text-sm font-bold text-white">← Back</button>
        <div className="flex items-center gap-3">
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
          {classWorks.map(w=><WorkCard key={w.id} w={w} users={users} evals={evals} comps={comps} likes={likes} setLikes={setLikes} setEvals={setEvals} setWorks={setWorks} certs={certs} onPromote={onPromote}/>)}
        </div>
      </main>
    </div>
  );
}

/* ════════════════════════════════════
   4. GLOBAL WORKS
   ════════════════════════════════════ */
function GlobalWorksPage({users,works,evals,comps,likes,setLikes,setEvals,setWorks,certs,onPromote,onBack}:{users:User[];works:Work[];evals:Ev[];comps:Comp[];likes:Like[];setLikes:React.Dispatch<React.SetStateAction<Like[]>>;setEvals:React.Dispatch<React.SetStateAction<Ev[]>>;setWorks:React.Dispatch<React.SetStateAction<Work[]>>;certs:Cert[];onPromote:(wid:string,cid:string,code:string)=>boolean;onBack:()=>void}){
  const sorted=useMemo(()=>[...works].sort((a,b)=>b.at.localeCompare(a.at)),[works]);
  return(
    <div className="relative min-h-screen bg-[#0d1230] text-white"><Bg/>
      <header className="sticky top-0 z-20 border-b border-white/10 bg-[#0f1535]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center">
        <button onClick={onBack} className="rounded-full border border-white/10 bg-white/5 px-5 py-2.5 text-sm font-bold text-white">← Back</button>
        <div className="flex items-center gap-3">
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
          {sorted.map(w=><WorkCard key={w.id} w={w} users={users} evals={evals} comps={comps} likes={likes} setLikes={setLikes} setEvals={setEvals} setWorks={setWorks} certs={certs} onPromote={onPromote}/>)}
        </div>
      </main>
    </div>
  );
}

/* ════════════════════════════════════
   5. WINNERS PAGE
   ════════════════════════════════════ */
function WinnersPage({comps,works,evals,users,likes,setLikes,certs,onPromote,adminPw,onArchive,onBack}:{comps:Comp[];works:Work[];evals:Ev[];users:User[];likes:Like[];setLikes:React.Dispatch<React.SetStateAction<Like[]>>;certs:Cert[];onPromote:(wid:string,cid:string,code:string)=>boolean;adminPw:string;onArchive:(comp:Comp,winners:any[])=>void;onBack:()=>void}){
  const today=new Date().toISOString().slice(0,10);
  const closedComps=useMemo(()=>comps.filter(c=>c.closingDate<=today&&!c.archived),[comps,today]);
  const[selCompId,setSelCompId]=useState<string|null>(closedComps[0]?.id||null);
  const[archivePw,setArchivePw]=useState("");
  const[archiveErr,setArchiveErr]=useState("");
  const[archiveMode,setArchiveMode]=useState(false);
  const[archiveDone,setArchiveDone]=useState<string|null>(null);
  const[viewEvalId,setViewEvalId]=useState<string|null>(null);
  const viewEv=evals.find(e=>e.id===viewEvalId)||null;

  const selComp=comps.find(c=>c.id===selCompId)||null;
  const isWinnerFn=(wid:string)=>certs.some(c=>c.workId===wid);
  void onPromote;

  const rankingsForComp=useMemo(()=>{
    if(!selComp)return[];
    const scored=works.filter(w=>w.competitionId===selComp.id&&isApproved(w)&&evals.some(e=>e.workId===w.id)).map(w=>{
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

      <section className="relative z-10 px-6 py-14 text-center border-b border-white/10 bg-[radial-gradient(ellipse_at_top,rgba(253,181,31,0.15),transparent_45%)]">
        <div className="text-5xl mb-4">🏅</div>
        <h1 className="text-4xl sm:text-6xl font-black tracking-tight">Hall of <span className="text-transparent bg-clip-text bg-gradient-to-r from-[#FDB51F] via-yellow-300 to-[#FDB51F]">Champions</span></h1>
        <p className="mt-3 text-xl font-bold text-white/50" dir="rtl">قاعة الفائزين — أيقونات المسابقات المغلقة</p>
        <p className="text-sm text-white/30 mt-2">Only closed competitions appear here</p>
      </section>

      <main className="max-w-6xl mx-auto px-6 py-12">
        {closedComps.length===0&&(
          <div className="rounded-3xl border border-white/10 bg-white/5 p-16 text-center">
            <div className="text-5xl mb-4">⏳</div>
            <p className="text-xl text-white/50">No closed competitions yet</p>
          </div>
        )}

        {closedComps.length>0&&(
          <>
            <div className="mb-10">
              <h2 className="text-xl font-black text-white/70 mb-5 uppercase tracking-widest">Select a Closed Competition</h2>
              <div className="flex flex-wrap gap-4">
                {closedComps.map(c=>(
                  <button key={c.id} onClick={()=>{setSelCompId(c.id);setArchiveMode(false);setArchiveDone(null)}}
                    className={cn("group relative rounded-3xl border p-5 flex flex-col items-center gap-3 w-44 transition-all hover:-translate-y-1",selCompId===c.id?"border-[#FDB51F]/50 bg-[#FDB51F]/15 shadow-lg shadow-[#FDB51F]/10":"border-white/10 bg-white/5 hover:border-[#FDB51F]/30")}>
                    <div className="absolute -top-2 -right-2 bg-rose-500 text-white text-[9px] font-black px-2 py-0.5 rounded-full uppercase">Closed</div>
                    <img src={c.logoUrl||"/gems-logo.png"} alt="" className="w-16 h-16 rounded-2xl bg-white object-contain p-1.5 shadow"/>
                    <div className="text-center">
                      <p className="text-sm font-black text-white leading-tight">{c.nameEn}</p>
                    </div>
                    {selCompId===c.id&&<div className="w-2 h-2 rounded-full bg-[#FDB51F]"/>}
                  </button>
                ))}
              </div>
            </div>

            {selComp&&(
              <div>
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-8 p-6 rounded-[2rem] border border-[#FDB51F]/20 bg-[#FDB51F]/8 backdrop-blur-sm">
                  <div className="flex items-center gap-4">
                    <img src={selComp.logoUrl||"/gems-logo.png"} alt="" className="h-16 w-16 rounded-2xl bg-white object-contain p-1.5"/>
                    <div><h3 className="text-2xl font-black text-white">{selComp.nameEn}</h3><div className="flex gap-2 mt-2"><Pill t="pink">Closed: {selComp.closingDate}</Pill><Pill t="gold">{rankingsForComp.length} Ranked</Pill></div></div>
                  </div>
                  <button onClick={()=>setArchiveMode(true)} className="rounded-2xl bg-white/10 border border-white/20 px-5 py-3 text-sm font-black text-white hover:bg-white/15 transition flex items-center gap-2 shrink-0">
                    🗄️ Send to Archive
                  </button>
                </div>

                {archiveMode&&(
                  <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 p-4 backdrop-blur-sm">
                    <div className="w-full max-w-sm rounded-3xl border border-white/10 bg-[#0f1535] p-6 shadow-2xl">
                      <h3 className="text-xl font-black text-white mb-2">🗄️ Archive Competition</h3>
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

                {archiveDone&&<div className="mb-6 rounded-2xl border border-emerald-400/20 bg-emerald-500/10 px-5 py-3 text-emerald-200 font-bold">✅ Archived successfully!</div>}

                <div className="space-y-6">
                  {rankingsForComp.map(({work,score,stu,rank})=>{
                    const wl=likes.filter(l=>l.workId===work.id);
                    const isTop3=rank<=3;
                    return(
                      <article key={work.id} className={cn("rounded-[2rem] border overflow-hidden",rank===1?"border-[#FDB51F]/40 bg-gradient-to-br from-[#FDB51F]/10 via-[#1F2E76]/50 to-[#0d1230] shadow-lg shadow-[#FDB51F]/10":rank===2?"border-slate-400/25 bg-[#1F2E76]/30":rank===3?"border-amber-700/25 bg-[#1F2E76]/25":"border-white/10 bg-white/5")}>
                        <div className="grid lg:grid-cols-[380px_1fr]">
                            <div className="relative"><div className="absolute top-4 left-4 z-10 flex items-center gap-2"><div className={cn("rounded-2xl px-4 py-2 text-2xl font-black shadow-lg backdrop-blur-md",rank===1?"bg-[#FDB51F] text-slate-950":rank===2?"bg-slate-300 text-slate-950":rank===3?"bg-amber-700 text-white":"bg-white/10 text-white border border-white/15")}>{MEDALS[rank-1]||`#${rank}`}</div>{isTop3&&<div className="rounded-full bg-white/10 backdrop-blur px-3 py-1 text-xs font-black text-white/80">#{rank} Champion</div>}</div><SecureVideo path={work.videoUrl} poster={work.thumb} controls className="w-full h-full object-cover min-h-[200px]"/></div>
                          <div className="p-6 flex flex-col justify-between">
                            <div><h3 className="text-xl font-black text-white">{work.title}</h3><div className="mt-3 flex items-center gap-3"><div className={cn("flex h-12 w-12 items-center justify-center rounded-2xl text-lg font-black shrink-0",rank===1?"bg-[#FDB51F] text-slate-900":rank===2?"bg-slate-300 text-slate-900":rank===3?"bg-amber-700 text-white":"bg-[#1F2E76] text-[#FDB51F]")}>{stu!.fullName[0]}</div><div><p className="text-base font-black text-white">{stu!.fullName} {isWinnerFn(work.id)&&"🌟"}</p><p className="text-sm text-white/40">{stu!.grade} • Section {stu!.section}</p></div></div></div>
                            <div className="mt-4 space-y-3"><div className="flex items-center gap-3"><div className={cn("rounded-2xl px-5 py-3 flex items-center gap-3",rank===1?"bg-[#FDB51F]/15 border border-[#FDB51F]/30":"bg-emerald-500/10 border border-emerald-400/20")}><span className="text-xs font-bold text-white/50">Score</span><span className={cn("text-3xl font-black",rank===1?"text-[#FDB51F]":"text-emerald-300")}>{score}</span><span className="text-white/30">/100</span></div><button onClick={()=>setViewEvalId(evals.find(e=>e.workId===work.id)?.id || null)} className="rounded-2xl bg-white/5 border border-white/10 px-4 py-3 text-xs font-bold text-white/60 hover:bg-white/10 transition">📊 Details</button></div>
                            <div className="flex flex-wrap gap-3 text-xs font-bold"><span className="text-sky-300">👫 {wl.filter(l=>l.type==="friend").length}</span><span className="text-emerald-300">👨‍👩‍👧 {wl.filter(l=>l.type==="parent").length}</span><span className="text-pink-300">👩‍🏫 {wl.filter(l=>l.type==="teacher").length}</span><span className="text-white/30">❤️ {wl.length} Total</span></div>
                            <LikesPanel workId={work.id} likes={likes} setLikes={setLikes}/></div>
                          </div>
                        </div>
                      </article>
                    );
                  })}
                </div>
              </div>
            )}
          </>
        )}
      </main>
      {viewEv&&(
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 p-4 backdrop-blur-sm" onClick={()=>setViewEvalId(null)}>
          <div className="w-full max-w-2xl rounded-[2rem] border border-white/10 bg-[#0f1535] p-6 shadow-2xl max-h-[90vh] overflow-y-auto" onClick={e=>e.stopPropagation()}>
            <div className="flex items-center justify-between mb-6"><div><h3 className="text-2xl font-black text-white">Scorecard</h3></div><button onClick={()=>setViewEvalId(null)} className="rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm font-bold text-white hover:bg-white/10">✕</button></div>
            <div className="space-y-3">{DEFAULT_CRITERIA.map((cr,i)=>(<div key={i} className="flex items-center justify-between rounded-2xl border border-white/10 bg-white/5 px-4 py-3"><span className="text-sm font-bold text-white/75">{cr}</span><span className="text-lg font-black text-sky-300">{viewEv.scores[i]}/10</span></div>))}</div>
            <div className="mt-6 flex items-center justify-between rounded-2xl border border-[#FDB51F]/20 bg-[#FDB51F]/10 px-5 py-4"><span className="text-sm font-bold text-[#FFE29A]">Final</span><span className="text-3xl font-black text-white">{tot(viewEv.scores)}/100</span></div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ════════════════════════════════════
   6. AUTH PAGES
   ════════════════════════════════════ */
function AdminLoginPage({adminPw,onBack,onOk}:{adminPw:string;onBack:()=>void;onOk:()=>void}){
  const[pw,setPw]=useState(""); const[err,setErr]=useState(""); const[showPw,setShowPw]=useState(false);

  /* ── forgot password flow (hidden by default) ── */
  const[forgot,setForgot]=useState(false);
  const[email,setEmail]=useState(""); const[fpErr,setFpErr]=useState(""); const[fpOk,setFpOk]=useState(false);
  const[attempts,setAttempts]=useState(0); const[revealed,setRevealed]=useState(false); const[busy,setBusy]=useState(false);
  const MAX_TRIES=5;
  const[mailStatus,setMailStatus]=useState<"idle"|"sending"|"sent"|"manual"|"failed"|"blocked">("idle");

  function sub(e:React.FormEvent){e.preventDefault(); const t=pw.trim(); if(t===adminPw){setErr("");onOk()}else{setErr("Incorrect password");setPw("")}}

  function openMailClient(toEmail:string){
    const body=[
      "GWA Champions — Administrator access",
      "بيانات دخول بوابة الإدارة",
      "",
      `Password: ${adminPw}`,
      `Recovery code: ${ADMIN_RECOVERY_CODE}`,
      "",
      "Portal: https://gwa-champions-platform.vercel.app",
      "بعد الدخول يمكنك تغيير كلمة المرور من تبويب Security.",
    ].join("\n");
    window.location.href=`mailto:${toEmail}?subject=${encodeURIComponent("GWA Champions — Admin access details")}&body=${encodeURIComponent(body)}`;
  }

  async function sendRecovery(toEmail:string){
    setMailStatus("sending");
    try{ await navigator.clipboard.writeText(`Password: ${adminPw}\nRecovery code: ${ADMIN_RECOVERY_CODE}`); }catch{}
    try{
      const res=await fetch("/api/send-mail",{
        method:"POST",
        headers:{"Content-Type":"application/json"},
        body:JSON.stringify({to:toEmail,password:adminPw,code:ADMIN_RECOVERY_CODE}),
      });
      const data=await res.json().catch(()=>({}));
      if(res.ok && data?.ok){
        setMailStatus("sent");
        return;
      }
      if(data?.code==="recipient-not-allowed"){ setMailStatus("blocked"); return; }
      if(res.status===502 || res.status===403){ setMailStatus("failed"); return; }
    }catch{}
    /* no mail provider configured (or network issue) → honest manual fallback */
    setMailStatus("manual");
  }

  function submitForgot(e:React.FormEvent){
    e.preventDefault();
    if(fpOk) return;
    const value=email.trim().toLowerCase();
    if(!value){setFpErr("أدخل البريد الإلكتروني • Enter your email");return}
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value)){setFpErr("صيغة البريد الإلكتروني غير صحيحة • Invalid email format");return}
    setBusy(true);
    const correct = isAdminEmail(value);
    if(correct){
      setFpErr(""); setFpOk(true);
      void sendRecovery(value);
      setBusy(false);
      return;
    }
    const used = attempts + 1;
    setAttempts(used);
    if(used >= MAX_TRIES){
      setRevealed(true);
      setFpErr("تجاوزت عدد المحاولات المسموح • Too many attempts");
    }else{
      setFpErr(`البريد الإلكتروني غير مسجّل • Email not registered — المحاولة ${used} من ${MAX_TRIES}`);
    }
    setBusy(false);
  }

  return(
    <div className="relative flex min-h-screen items-center justify-center overflow-hidden bg-[#0d1230] px-6 text-white"><Bg/>
      <div className="relative w-full max-w-md rounded-[2.4rem] border border-white/10 bg-white/5 p-8 shadow-2xl backdrop-blur-xl">
        <button onClick={onBack} className="mb-6 text-sm font-bold text-white/50 hover:text-white">← Back</button>
        <div className="mb-8 flex flex-col items-center"><Logo/><h2 className="mt-5 text-2xl font-black text-white">Administration Portal</h2><p className="mt-1 text-sm text-white/50">بوابة الإدارة</p></div>

        <form onSubmit={sub} className="space-y-5">
          <div>
            <FL>Admin Password</FL>
            <div className="relative">
              <Inp type={showPw?"text":"password"} value={pw} onChange={e=>{setPw(e.target.value);if(err)setErr("")}} className="pr-14" autoComplete="off"/>
              <button type="button" onClick={()=>setShowPw(s=>!s)} className="absolute right-4 top-1/2 -translate-y-1/2 text-xs font-bold text-white/40">{showPw?"Hide":"Show"}</button>
            </div>
          </div>
          {err&&<div className="rounded-2xl border border-rose-400/20 bg-rose-500/10 px-4 py-3 text-center text-sm font-bold text-rose-200">⚠️ {err}</div>}
          <button type="submit" className="w-full rounded-2xl bg-[#FDB51F] px-5 py-4 text-base font-black text-slate-950 shadow-lg shadow-[#FDB51F]/20 hover:bg-[#F7C84A]">دخول الإدارة • Enter</button>
        </form>

        <div className="mt-4 text-center">
          <button type="button" onClick={()=>{setForgot(true);setFpErr("");setFpOk(false);setEmail("");setMailStatus("idle");}} className="text-[12px] font-bold text-white/40 underline decoration-dotted transition hover:text-[#FFE29A]">Forgot password? • نسيت كلمة المرور</button>
        </div>
      </div>

      {forgot&&(
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/85 p-4 backdrop-blur-md" onClick={()=>setForgot(false)}>
          <div className="w-full max-w-md overflow-hidden rounded-[2rem] border border-white/10 bg-[#0f1535] shadow-2xl" onClick={e=>e.stopPropagation()}>
            <div className="h-1.5 w-full bg-gradient-to-r from-[#1697E8] via-[#FDB51F] to-[#10B09B]"/>
            <div className="p-6">
              <div className="mb-5 text-center">
                <div className="mx-auto mb-3 grid h-14 w-14 place-items-center rounded-2xl bg-[#1697E8]/15 text-2xl ring-1 ring-[#1697E8]/25">🔑</div>
                <h3 className="text-lg font-black text-white">Forgot Password</h3>
                <p className="mt-1 text-sm text-white/50">استعادة كلمة مرور الإدارة</p>
              </div>

              {!fpOk ? (
                <form onSubmit={submitForgot} className="space-y-3">
                  <p className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3 text-xs font-bold leading-6 text-white/55">
                    أدخل <span className="text-white">البريد الإلكتروني المسجّل للإدارة</span> وسنرسل إليك كلمة المرور ورمز الاستعادة.
                    <br/>Enter the registered admin email and we will send the credentials to it.
                  </p>
                  <div><FL>Registered email • البريد المسجّل</FL><Inp type="email" value={email} onChange={e=>{setEmail(e.target.value);setFpErr("")}} placeholder="name@gemsedu.com" autoFocus autoComplete="off"/></div>

                  {fpErr&&<div className="rounded-xl border border-rose-400/20 bg-rose-500/10 px-4 py-3 text-sm font-bold leading-6 text-rose-200">{fpErr}</div>}

                  {revealed&&(
                    <div className="rounded-xl border border-[#FDB51F]/30 bg-[#FDB51F]/10 px-4 py-3 text-sm font-bold leading-7 text-[#FFE29A]">
                      ✅ البريد المسجّل هو:<br/>
                      <span dir="ltr" className="block font-black tracking-wide text-white">{ADMIN_EMAIL}</span>
                      <span dir="ltr" className="mt-1 block text-[11px] font-bold text-[#FFE29A]/80">أو • or: {ADMIN_EMAIL_ALT}</span>
                      <button type="button" onClick={()=>{setEmail(ADMIN_EMAIL);setFpErr("");}} className="mt-2 block w-full rounded-xl bg-[#FDB51F] py-2.5 text-xs font-black text-slate-950 hover:bg-[#F7C84A]">استخدام {ADMIN_EMAIL}</button>
                      <button type="button" onClick={()=>{setEmail(ADMIN_EMAIL_ALT);setFpErr("");}} className="mt-1.5 block w-full rounded-xl border border-[#FDB51F]/40 bg-white/5 py-2.5 text-xs font-black text-[#FFE29A] hover:bg-white/10">استخدام {ADMIN_EMAIL_ALT}</button>
                    </div>
                  )}

                  {!revealed&&attempts>0&&(
                    <p className="text-center text-[11px] font-bold text-white/35">المحاولات المتبقية: {MAX_TRIES-attempts}</p>
                  )}

                  <div className="flex gap-3 pt-1">
                    <button type="button" onClick={()=>setForgot(false)} className="flex-1 rounded-2xl border border-white/10 bg-white/5 py-3 text-sm font-black text-white/70 hover:bg-white/10">Cancel</button>
                    <button type="submit" disabled={busy} className={cn("flex-[1.4] rounded-2xl py-3 text-sm font-black transition",busy?"cursor-not-allowed bg-white/10 text-white/30":"bg-[#1697E8] text-white hover:bg-sky-400")}>{busy?"Sending…":"Send details 📧"}</button>
                  </div>
                </form>
              ) : (
                <div className="space-y-3">
                  {mailStatus==="sending"&&(
                    <div className="rounded-2xl border border-sky-400/25 bg-sky-500/10 p-4 text-center">
                      <p className="text-2xl">⏳</p>
                      <p className="mt-2 text-sm font-black text-sky-200">جاري الإرسال… • Sending</p>
                    </div>
                  )}
                  {mailStatus==="sent"&&(
                    <div className="rounded-2xl border border-emerald-400/25 bg-emerald-500/10 p-4 text-center">
                      <p className="text-2xl">✅</p>
                      <p className="mt-2 text-sm font-black text-emerald-200">تم إرسال البيانات فعلياً</p>
                      <p className="mt-1 text-[12px] leading-6 text-emerald-100/80" dir="rtl">أُرسلت كلمة المرور ورمز الاستعادة إلى بريد الإدارة المسجّل. تحقّق من صندوق الوارد (ومجلد الرسائل غير المرغوبة).</p>
                      <p className="mt-1 text-[12px] font-black text-white" dir="ltr">{ADMIN_EMAIL}</p>
                    </div>
                  )}
                  {mailStatus==="blocked"&&(
                    <div className="rounded-2xl border border-[#FDB51F]/30 bg-[#FDB51F]/10 p-4">
                      <p className="text-center text-2xl">⚠️</p>
                      <p className="mt-2 text-center text-sm font-black text-[#FFE29A]">الإرسال التلقائي محصور ببريد مالك حساب Resend</p>
                      <p className="mt-2 text-[12px] font-bold leading-6 text-white/65" dir="rtl">
                        المفتاح مربوط لكن Resend لا يسمح بإرسال تجريبي إلى هذا العنوان إلا بعد توثيق دومين المدرسة في resend.com/domains ثم ضبط MAIL_FROM على بريد من الدومين الموثّق.
                        <br/>يمكنك الآن استخدام زر الرسالة الجاهزة أدناه، أو توثيق الدومين ليعمل الإرسال تلقائياً.
                      </p>
                    </div>
                  )}
                  {mailStatus==="failed"&&(
                    <div className="rounded-2xl border border-rose-400/25 bg-rose-500/10 p-4">
                      <p className="text-center text-2xl">📭</p>
                      <p className="mt-2 text-center text-sm font-black text-rose-200">تعذّر إرسال البريد</p>
                      <p className="mt-1 text-center text-[12px] font-bold leading-6 text-rose-100/80" dir="rtl">تحقّق من صلاحية مفتاح Resend في Vercel ثم أعد المحاولة، أو استخدم الإرسال اليدوي.</p>
                    </div>
                  )}
                  {mailStatus==="manual"&&(
                    <div className="rounded-2xl border border-[#FDB51F]/30 bg-[#FDB51F]/10 p-4">
                      <p className="text-center text-2xl">✉️</p>
                      <p className="mt-2 text-center text-sm font-black text-[#FFE29A]">الإرسال التلقائي غير مُفعّل بعد</p>
                      <p className="mt-2 text-[12px] font-bold leading-6 text-white/65" dir="rtl">
                        لا توجد خدمة بريد مربوطة بالموقع، لذلك لم تُرسل الرسالة تلقائياً. اضغط الزر أدناه لتفتح رسالة جاهزة في برنامج البريد على جهازك — تحتوي كلمة المرور ورمز الاستعادة — ثم اضغط «إرسال». وقد نُسخت البيانات أيضاً إلى الحافظة.
                      </p>
                      <p className="mt-2 text-center text-[11px] font-black text-white" dir="ltr">{ADMIN_EMAIL}</p>
                    </div>
                  )}
                  <button type="button" onClick={()=>void sendRecovery(ADMIN_EMAIL)} disabled={mailStatus==="sending"} className={cn("w-full rounded-2xl py-3.5 text-sm font-black text-white transition",mailStatus==="sending"?"cursor-not-allowed bg-white/10 text-white/30":"bg-[#1697E8] hover:bg-sky-400")}>{mailStatus==="manual"?"📧 فتح رسالة جاهزة للإرسال":"↻ إعادة المحاولة"}</button>
                  <button type="button" onClick={()=>openMailClient(ADMIN_EMAIL)} className="w-full rounded-2xl border border-white/10 bg-white/5 py-3 text-xs font-black text-white/60 hover:bg-white/10">✉️ إرسال يدوي من بريدي • Manual send</button>
                  <button type="button" onClick={()=>{setForgot(false);setFpOk(false);}} className="w-full rounded-2xl border border-white/10 bg-white/5 py-3 text-sm font-black text-white/70 hover:bg-white/10">إغلاق • Close</button>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function StudentLoginPage({users,onBack,onOk}:{users:User[];onBack:()=>void;onOk:(s:User)=>void}){
  const[u,setU]=useState(""); const[p,setP]=useState(""); const[err,setErr]=useState(""); const[loading,setLoading]=useState(false);
  const norm=(x:any)=>String(x||"").trim().toLowerCase();
  const mapRow=(r:any):User=>({
    id:r.id, fullName:r.full_name||r.username||"Student", role:(r.role||"student") as Role,
    grade:r.grade||"", section:r.section||"", teacherId:r.teacher_id||undefined,
    username:r.username, password:String(r.password??""), createdAt:r.created_at||"",
    active:r.active!==false
  });

  /** Cloud lookup — the source of truth, so the same login works on every device */
  async function lookupCloud(username:string){
    if(!isSupabaseConfigured) return {rows:[] as User[], error:"not-configured"};
    const {data,error}=await supabase.from("accounts").select("*").ilike("username",username).limit(20);
    if(error) return {rows:[] as User[], error:error.message};
    const rows=((data||[]) as any[]).map(mapRow);
    if(rows.length>0){
      try{
        const raw=localStorage.getItem(SK.u);
        const cached:User[]=raw?JSON.parse(raw):[];
        const merged=[...rows,...cached.filter(c=>!rows.some(x=>x.id===c.id))];
        localStorage.setItem(SK.u,JSON.stringify(merged));
      }catch{}
    }
    return {rows, error:""};
  }

  async function sub(e:React.FormEvent){
    e.preventDefault();
    setErr("");
    const username=u.trim(); const password=p;
    if(!username||!password){setErr("أدخل اسم المستخدم وكلمة المرور • Enter username and password");return}
    setLoading(true);
    try{
      // 1) Cached/synced list (works offline and on repeat visits)
      let s=users.find(x=>norm(x.username)===norm(username) && String(x.password)===password && (x.role==="student"||x.role==="teacher") && x.active!==false);

      // 2) Direct cloud check (fresh device, or account created on another device)
      if(!s){
        const cloud=await lookupCloud(username);
        if(cloud.error && cloud.error!=="not-configured") console.warn("accounts lookup:", cloud.error);
        if(cloud.rows.length>0){
          const samePassword=cloud.rows.find(r=>String(r.password)===password);
          if(samePassword && (samePassword.role==="student"||samePassword.role==="teacher") && samePassword.active!==false){
            s=samePassword;
          }else if(samePassword){
            setErr("هذا الحساب غير مفعّل • This account is inactive");
            setLoading(false);return;
          }else{
            setErr("كلمة المرور غير صحيحة • Incorrect password");
            setLoading(false);return;
          }
        }
      }

      if(s){
        try{
          const raw=localStorage.getItem(SK.u);
          const cached:User[]=raw?JSON.parse(raw):[];
          if(!cached.some(c=>c.id===s!.id)) localStorage.setItem(SK.u,JSON.stringify([s,...cached]));
        }catch{}
        onOk(s); return;
      }
      setErr("اسم المستخدم غير موجود • Username not found — تأكد من الاسم أو اطلبه من المعلم");
    }catch(e:any){
      setErr("تعذر تسجيل الدخول • Login error: "+(e?.message||"unknown"));
    }finally{ setLoading(false); }
  }

  return(
    <div className="relative min-h-screen bg-[#0d1230] flex items-center justify-center px-6"><Bg/>
      <div className="relative w-full max-w-md rounded-[2.4rem] border border-white/10 bg-white/5 p-8 backdrop-blur-xl shadow-2xl"><button onClick={onBack} className="text-white/50 text-sm font-bold mb-6 hover:text-white">← Back</button><div className="flex flex-col items-center mb-8"><Logo/><h2 className="mt-5 text-2xl font-black text-white">Student Login • دخول الطالب</h2><p className="text-xs text-white/40 mt-2 text-center">Cloud accounts — same login on every device</p><p className="text-[11px] text-[#FDB51F] mt-2 text-center font-bold">الطلاب والمعلمون يدخلون من هنا • Students &amp; teachers sign in here</p></div>
        <form onSubmit={sub} className="space-y-5"><div><FL>Username • اسم المستخدم</FL><Inp value={u} onChange={e=>setU(e.target.value)} placeholder="mariam.salem" autoComplete="username"/></div><div><FL>Password • كلمة المرور</FL><Inp type="password" value={p} onChange={e=>setP(e.target.value)} autoComplete="current-password"/></div>
          {err&&<div className="rounded-2xl border border-rose-400/20 bg-rose-500/10 px-4 py-3 text-sm text-rose-200 font-bold">{err}</div>}
          <button type="submit" disabled={loading} className={cn("w-full rounded-2xl px-5 py-4 text-base font-black shadow-lg shadow-[#FDB51F]/20",loading?"bg-white/10 text-white/30 cursor-not-allowed":"bg-[#FDB51F] text-slate-950 hover:bg-[#F7C84A]")}>{loading?"Signing in...":"Login 🎓"}</button></form></div>
    </div>
  );
}

/* ════════════════════════════════════
   7. ADMIN PAGE
   ════════════════════════════════════ */

/* ── ADMIN: edit the classes a teacher teaches ── */
function TeacherClassesModal({teacher,onSave,onClose}:{teacher:User;onSave:(list:ClassRef[])=>void;onClose:()=>void}){
  const[list,setList]=useState<ClassRef[]>(()=>{const d=decodeClasses(teacher);return d.length?d:[{grade:"Grade 6",section:"A"}]});
  const[g,setG]=useState("Grade 6"); const[sec,setSec]=useState("A");
  function add(){ setList(prev=>{ if(prev.some(c=>c.grade===g&&c.section===sec)) return prev; return [...prev,{grade:g,section:sec}]; }); }
  return(
    <div className="fixed inset-0 z-[60] flex items-center justify-center bg-black/85 p-4 backdrop-blur-md" onClick={onClose}>
      <div className="w-full max-w-lg overflow-hidden rounded-[2rem] border border-white/10 bg-[#0f1535] shadow-2xl" onClick={e=>e.stopPropagation()}>
        <div className="h-1.5 w-full bg-gradient-to-r from-[#1697E8] via-[#10B09B] to-[#FDB51F]"/>
        <div className="p-6">
          <div className="mb-4 flex items-start justify-between gap-3">
            <div><h3 className="text-lg font-black text-white">🏫 Classes • الصفوف</h3><p className="text-sm text-white/45">{teacher.fullName} — @{teacher.username}</p></div>
            <button onClick={onClose} className="rounded-full border border-white/10 bg-white/5 px-4 py-2 text-sm font-bold text-white hover:bg-white/10">✕</button>
          </div>
          <div className="mb-4 flex flex-wrap gap-2">
            {list.map(c=>(
              <span key={`${c.grade}-${c.section}`} className="inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/5 px-3 py-1.5 text-xs font-bold text-white/80">
                {classLabel(c)}
                <button onClick={()=>setList(prev=>prev.filter(x=>!(x.grade===c.grade&&x.section===c.section)))} className="text-rose-300 hover:text-rose-200">✕</button>
              </span>
            ))}
            {list.length===0&&<span className="text-xs font-bold text-white/35">لم تُحدَّد صفوف بعد • No classes yet</span>}
          </div>
          <div className="grid grid-cols-[1fr_auto_auto] items-end gap-2">
            <div><FL>Grade</FL><Sel value={g} onChange={e=>setG(e.target.value)}>{GRADES.map(x=><option key={x} value={x}>{x}</option>)}</Sel></div>
            <div><FL>Section</FL><Sel value={sec} onChange={e=>setSec(e.target.value)}>{_SECTIONS.map(x=><option key={x} value={x}>{x}</option>)}</Sel></div>
            <button type="button" onClick={add} className="rounded-2xl bg-[#1697E8] px-4 py-3 text-sm font-black text-white hover:bg-sky-400">＋ Add</button>
          </div>
          <div className="mt-5 flex gap-3">
            <button onClick={onClose} className="flex-1 rounded-2xl border border-white/10 bg-white/5 py-3 text-sm font-black text-white/70">Cancel</button>
            <button onClick={()=>{onSave(list);onClose();}} className="flex-[1.4] rounded-2xl bg-[#FDB51F] py-3 text-sm font-black text-slate-950 hover:bg-[#F7C84A]">Save Classes ✓</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function AdminPage({users,setUsers,comps,setComps,archive,adminPw,setAdminPw,onBack}:{users:User[];setUsers:React.Dispatch<React.SetStateAction<User[]>>;comps:Comp[];setComps:React.Dispatch<React.SetStateAction<Comp[]>>;archive:ArchiveEntry[];adminPw:string;setAdminPw:(pw:string)=>void;onBack:()=>void}){
  const teachers=useMemo(()=>users.filter(u=>u.role==="teacher"),[users]);
  const[tab,setTab]=useState<"users"|"export"|"comps"|"archive"|"security">("users");
  const fRef=useRef<HTMLInputElement>(null);
  const[name,setName]=useState(""); const[role,setRole]=useState<Role>("student"); const[grade,setGrade]=useState("Grade 6"); const[sec,setSec]=useState("X"); const[tid,setTid]=useState(""); const[umsg,setUmsg]=useState("");
  const[nAr,setNAr]=useState(""); const[nEn,setNEn]=useState(""); const[cGrades,setCGrades]=useState<string[]>(["Grade 6"]); const[rAr,setRAr]=useState(()=>ruleTemplateAr("المسابقة")); const[rEn,setREn]=useState(()=>ruleTemplateEn("Competition")); const[cClose,setCClose]=useState("2026-06-01"); const[cCert,setCCert]=useState("2026-06-15"); const[cmsg,setCmsg]=useState(""); const[impMsg,setImpMsg]=useState("");
  const[compCriteria,setCompCriteria]=useState<string[]>([...DEFAULT_CRITERIA]);
  const[selPreset,setSelPreset]=useState<string>("");
  const[editClassesFor,setEditClassesFor]=useState<User|null>(null);
  /* teacher class builder state */
  const[tGrade,setTGrade]=useState("Grade 6"); const[tSec,setTSec]=useState("A");
  const[tClasses,setTClasses]=useState<ClassRef[]>([{grade:"Grade 6",section:"A"}]);
  /* security tab state */
  const[pwOld,setPwOld]=useState(""); const[pwNew,setPwNew]=useState(""); const[pwConfirm,setPwConfirm]=useState("");
  const[pwMsg,setPwMsg]=useState(""); const[pwErr,setPwErr]=useState(""); const[showCode,setShowCode]=useState(false);

  function applyPreset(idx:string){
    setSelPreset(idx);
    if(idx==="custom"){
      setNAr("");setNEn("");
      setCompCriteria([...DEFAULT_CRITERIA]);
      const t=rulesForPreset(undefined,"المسابقة","Competition");
      setRAr(t.ar); setREn(t.en);
      return;
    }
    const p=COMP_PRESETS[Number(idx)];
    if(p){
      setNAr(p.nameAr);setNEn(p.nameEn);setCompCriteria([...p.criteria]);
      const t=rulesForPreset(Number(idx));
      setRAr(t.ar); setREn(t.en);
    }
  }

  function createUser(e:React.FormEvent){
    e.preventDefault();
    if(!name.trim()){setUmsg("Enter name");return}
    if(role==="student"&&!tid){setUmsg("Select teacher");return}
    if(role==="teacher"&&tClasses.length===0){setUmsg("أضف صفاً واحداً على الأقل للمعلم • Add at least one class");return}
    const u=mkUser(name,role,grade,sec,role==="student"?tid:undefined,users);
    if(role==="teacher"){
      u.grade=encodeClasses(tClasses);
      u.section="";
    }
    setUsers(p=>[u,...p]);
    setUmsg(`✅ ${u.username} / ${u.password}${role==="teacher"?` • ${tClasses.length} classes`:""}`);
    setName("");
  }
  function saveTeacherClasses(list:ClassRef[]){
    if(!editClassesFor) return;
    const id=editClassesFor.id;
    setUsers(prev=>prev.map(it=>it.id===id?{...it,grade:encodeClasses(list),section:""}:it));
  }
  function createComp(e:React.FormEvent){e.preventDefault();if(!nAr.trim()||!nEn.trim()){setCmsg("Enter both names");return}
    const code=Math.random().toString(36).slice(2,8).toUpperCase();
    const t=rulesForPreset(selPreset&&selPreset!=="custom"?Number(selPreset):undefined,nAr,nEn);
    const finalAr=(rAr.trim()||t.ar);
    const finalEn=(rEn.trim()||t.en);
    setComps(p=>[{id:mid("c"),nameAr:nAr,nameEn:nEn,logoUrl:"/gems-logo.png",grades:cGrades,rulesAr:finalAr,rulesEn:finalEn,closingDate:cClose,certDate:cCert,started:true,archived:false,criteria:compCriteria,evalCode:code},...p]);
    setCmsg(`✅ Created • Evaluation Code: ${code}`);
    setNAr("");setNEn("");setCompCriteria([...DEFAULT_CRITERIA]);setRAr(ruleTemplateAr("المسابقة"));setREn(ruleTemplateEn("Competition"))}
  function findCol(h:string[],p:string[]){for(let i=0;i<h.length;i++){const s=h[i].toLowerCase().replace(/[\s_\-().]+/g,"");for(const x of p)if(s===x||s.includes(x))return i}return -1}
  async function handleExcelUpload(e:React.ChangeEvent<HTMLInputElement>){const file=e.target.files?.[0];if(!file)return;
    try{const buf=await file.arrayBuffer();const wb=XLSX.read(buf,{type:"array"});const raw=XLSX.utils.sheet_to_json<(string|number)[]>(wb.Sheets[wb.SheetNames[0]],{header:1,defval:""});if(raw.length<2){setImpMsg("⚠️ Empty file");return}
      const headers=raw[0].map(h=>String(h));const rows=raw.slice(1);const ni=findCol(headers,["name","fullname","الاسم"]),ri=findCol(headers,["role","الدور"]),gi=findCol(headers,["grade","الصف"]),si=findCol(headers,["section","القسم","sec"]);const ai=ni>=0?ni:0;const imported:User[]=[];
      rows.forEach(r=>{const n=String(r[ai]||"").trim();if(!n||n.toLowerCase()==="name")return;const rl:Role=["teacher","معلم"].includes(String(r[ri]||"").toLowerCase())?"teacher":"student";imported.push(mkUser(n,rl,String(r[gi]||"Grade 6").trim(),String(r[si]||"X").trim(),undefined,[...users,...imported]))});
      if(!imported.length){setImpMsg("⚠️ No valid rows");return}setUsers(p=>[...imported,...p]);setImpMsg(`✅ Imported ${imported.length} users`);e.target.value=""}catch{setImpMsg("⚠️ Error reading Excel")}}
  function downloadExcel(){const data=users.map(u=>({Name:u.fullName,Role:u.role,Grade:u.grade,Section:u.section,Username:u.username,Password:u.password,Active:u.active}));const ws=XLSX.utils.json_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Users");XLSX.writeFile(wb,"gwa-users.xlsx")}
  function downloadTemplate(){const t=[{Name:"Mariam Salem",Role:"Student",Grade:"Grade 6",Section:"X"}];const ws=XLSX.utils.json_to_sheet(t);ws["!cols"]=[{wch:25},{wch:10},{wch:12},{wch:10}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Users");XLSX.writeFile(wb,"gwa-template.xlsx")}
  
  function logoToBase64():Promise<string>{
    return new Promise((resolve)=>{
      const img=new Image();
      img.crossOrigin="anonymous";
      img.onload=()=>{
        const canvas=document.createElement("canvas");
        canvas.width=img.naturalWidth;
        canvas.height=img.naturalHeight;
        const ctx=canvas.getContext("2d");
        if(ctx){ctx.drawImage(img,0,0);resolve(canvas.toDataURL("image/png"))}
        else resolve("");
      };
      img.onerror=()=>resolve("");
      img.src="/gems-logo.png";
    });
  }

  async function generateCards(){
    const qrData:Record<string,string>={};for(const u of users){try{qrData[u.id]=await QRCode.toDataURL(`GWA|${u.fullName}|${u.username}|${u.password}`,{width:140,margin:1})}catch{qrData[u.id]=""}}
    const logoBase64=await logoToBase64();
    const win=window.open("","_blank");if(!win)return;
    const cards=users.map((u,i)=>{const isT=u.role==="teacher";const qr=qrData[u.id]||"";return`<div class="card"><div class="ci"><div class="ch"><div class="hl">${logoBase64?`<img src="${logoBase64}" class="logo" alt="GEMS"/>`:`<div class="logo-placeholder">GWA</div>`}<div><div class="pen">GWA Arabic & Islamic Champions</div><div class="par" style="color:${isT?"#1697E8":"#FDB51F"}">${isT?"Teacher":"Student"} Login</div></div></div>${qr?`<img src="${qr}" class="qr"/>`:""}</div><hr class="div"/><div class="inf"><div>Name : <b>${u.fullName}</b></div><div>ID : <b>${i+1}</b></div><div>Grade : <b>${u.grade}</b></div>${u.role==="student"?`<div>Section : <b>${u.section}</b></div>`:""}</div><div class="badge" style="background:${isT?"#1697E8":"#FDB51F"};color:${isT?"#fff":"#111"}">${u.role}</div><div class="creds"><div class="url">gwa.gems-world-academy.ae</div><div>• Username: <span class="hl2">${u.username}</span></div><div>• Password: <span class="hl2">${u.password}</span></div></div></div></div>`}).join("");
    win.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8"/><title>GWA Cards</title><style>*{box-sizing:border-box;margin:0;padding:0}body{background:#f0f0f0;font-family:Arial,sans-serif;padding:24px}.tb{display:flex;gap:12px;justify-content:center;margin-bottom:24px;padding:16px;background:#1F2E76;border-radius:16px}.tb button{padding:10px 24px;border-radius:999px;border:none;font-weight:700;cursor:pointer;background:#FDB51F;color:#111}.btn-close{background:rgba(255,255,255,0.15);color:#fff}.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:16px;max-width:1100px;margin:0 auto}.card{break-inside:avoid}.ci{background:#fff;border:3px solid #5ECEC8;border-radius:20px;padding:20px;min-height:220px}.ch{display:flex;justify-content:space-between;margin-bottom:12px}.hl{display:flex;align-items:center;gap:10px}.logo{width:60px;object-fit:contain;display:block}.logo-placeholder{width:52px;height:40px;background:#1F2E76;border-radius:8px;display:flex;align-items:center;justify-content:center;color:#FDB51F;font-weight:900;font-size:10px}.pen{font-size:14px;font-weight:800;color:#1F2E76}.par{font-size:11px;font-weight:700;margin-top:2px}.qr{width:90px}.div{border:none;border-top:1.5px solid #eee;margin:10px 0}.inf{font-size:13px;line-height:1.8}.badge{display:inline-block;border-radius:999px;padding:4px 14px;font-size:11px;font-weight:800;margin:10px 0}.creds{background:#fafafa;border-radius:10px;padding:10px;border:1px solid #eee;font-size:12px;line-height:1.9}.url{font-size:10px;color:#777;margin-bottom:4px}.hl2{background:#FDB51F;padding:1px 6px;border-radius:4px;font-weight:700}@media print{body{background:#fff;padding:8px}.tb{display:none}}</style></head><body><div class="tb"><button onclick="window.print()">🖨 Print</button><button onclick="window.print()">📄 Download PDF</button><button class="btn-close" onclick="window.close()">✕ Close</button></div><div class="grid">${cards}</div></body></html>`);win.document.close()}

  return(
    <div className="relative min-h-screen bg-[linear-gradient(180deg,#0d1230,#111a45,#0d1230)] text-white"><Bg/>
      <header className="sticky top-0 z-20 border-b border-white/10 bg-[#0f1535]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center">
        <div className="flex items-center gap-4"><Logo s/><div><div className="text-[10px] font-bold uppercase tracking-[0.28em] text-[#FDB51F]">Administration Portal • بوابة الإدارة</div><h1 className="text-xl font-black">User & Competition Management</h1></div></div>
        <div className="flex gap-3"><button onClick={()=>setTab("security")} className="rounded-full border border-white/10 bg-white/5 px-5 py-2 text-sm font-bold text-white hover:bg-white/10">🔑 Password</button><button onClick={onBack} className="rounded-full border border-white/10 bg-white/5 px-5 py-2 text-sm font-bold text-white hover:bg-white/10">Home</button></div>
      </header>
      <main className="max-w-6xl mx-auto px-6 py-8">
        <div className="flex flex-wrap gap-2 mb-8 bg-white/5 border border-white/10 rounded-2xl p-1.5 overflow-x-auto">
          {([["users","👥 Users"],["export","📊 Excel & Cards"],["comps","🏆 Competitions"],["archive","🗄️ Archive • الأرشيف"],["security","🔑 Security"]] as const).map(([k,l])=>(
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
                  <div className="grid grid-cols-2 gap-3">
                    <div><FL>Role</FL><Sel value={role} onChange={e=>setRole(e.target.value as Role)}><option value="student">Student</option><option value="teacher">Teacher</option></Sel></div>
                    {role==="student"&&<div><FL>Grade</FL><Sel value={grade} onChange={e=>setGrade(e.target.value)}>{GRADES.map(g=><option key={g} value={g}>{g}</option>)}</Sel></div>}
                    {role==="student"&&<div><FL>Section</FL><Sel value={sec} onChange={e=>setSec(e.target.value)}>{_SECTIONS.map(s=><option key={s} value={s}>{s}</option>)}</Sel></div>}
                  </div>
                  {role==="teacher"&&(
                    <div className="rounded-2xl border border-[#1697E8]/25 bg-[#1697E8]/10 p-4">
                      <p className="mb-2 text-xs font-black text-sky-200">🏫 الصفوف التي يدرّسها المعلم • Classes taught</p>
                      <div className="mb-3 flex flex-wrap gap-2">
                        {tClasses.map(c=>(
                          <span key={`${c.grade}-${c.section}`} className="inline-flex items-center gap-2 rounded-full border border-white/15 bg-black/25 px-3 py-1.5 text-xs font-bold text-white/85">
                            {classLabel(c)}
                            <button type="button" onClick={()=>setTClasses(prev=>prev.filter(x=>!(x.grade===c.grade&&x.section===c.section)))} className="text-rose-300 hover:text-rose-200">✕</button>
                          </span>
                        ))}
                        {tClasses.length===0&&<span className="text-xs font-bold text-white/35">لم تُضف صفوف بعد • No classes added</span>}
                      </div>
                      <div className="grid grid-cols-[1fr_auto_auto] items-end gap-2">
                        <div><FL>Grade</FL><Sel value={tGrade} onChange={e=>setTGrade(e.target.value)}>{GRADES.map(x=><option key={x} value={x}>{x}</option>)}</Sel></div>
                        <div><FL>Section</FL><Sel value={tSec} onChange={e=>setTSec(e.target.value)}>{_SECTIONS.map(x=><option key={x} value={x}>{x}</option>)}</Sel></div>
                        <button type="button" onClick={()=>setTClasses(prev=>prev.some(c=>c.grade===tGrade&&c.section===tSec)?prev:[...prev,{grade:tGrade,section:tSec}])} className="rounded-2xl bg-[#1697E8] px-4 py-3 text-sm font-black text-white hover:bg-sky-400">＋</button>
                      </div>
                    </div>
                  )}
                  {role==="student"&&<div><FL>Assigned Teacher</FL><Sel value={tid} onChange={e=>setTid(e.target.value)}><option value="">Select…</option>{teachers.map(t=><option key={t.id} value={t.id}>{t.fullName} — {t.grade}</option>)}</Sel></div>}
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
              <div className="space-y-3">{users.filter(u=>u.id!=="admin").map(u=>(
                <div key={u.id} className={cn("rounded-2xl border p-4 transition",u.active?"border-white/10 bg-black/20":"border-rose-500/20 bg-rose-500/5 opacity-60")}>
                  <div className="flex justify-between items-start gap-2">
                    <div className="min-w-0"><div className="font-black text-white flex items-center gap-2 flex-wrap">{u.fullName}{!u.active&&<span className="text-[9px] bg-rose-500 text-white px-2 py-0.5 rounded-full uppercase">Disabled</span>}</div><div className="flex flex-wrap gap-1.5 mt-2"><Pill t={u.role==="teacher"?"green":"blue"}>{u.role}</Pill>{u.role==="student"&&<Pill t="gold">{u.grade}</Pill>}{u.role==="student"&&<Pill t="pink">{u.section}</Pill>}{u.role==="teacher"&&decodeClasses(u).map(c=><Pill key={`${c.grade}-${c.section}`} t="green">🏫 {classLabel(c)}</Pill>)}{u.role==="student"&&u.teacherId&&<Pill t="slate">Teacher: {teachers.find(t=>t.id===u.teacherId)?.fullName||"—"}</Pill>}</div><div className="mt-1.5 text-xs text-white/40 font-bold">👤 {u.username} / 🔑 {u.password}</div></div>
                    <div className="flex flex-col gap-1.5 shrink-0">{u.role==="teacher"&&<button onClick={()=>setEditClassesFor(u)} className="rounded-xl px-3 py-1.5 text-[10px] font-black bg-sky-500/20 border border-sky-400/30 text-sky-200 hover:bg-sky-500/30 transition">🏫 Classes</button>}<button onClick={()=>setUsers(prev=>prev.map(it=>it.id===u.id?{...it,active:!it.active}:it))} className={cn("rounded-xl px-3 py-2 text-xs font-black transition",u.active?"bg-rose-500 text-white hover:bg-rose-600":"bg-emerald-500 text-white hover:bg-emerald-600")}>{u.active?"Disable":"Activate"}</button><button onClick={()=>{if(window.confirm(`Remove "${u.fullName}" permanently?`)){ setUsers(prev=>prev.filter(it=>it.id!==u.id)); }}} className="rounded-xl px-3 py-1.5 text-[10px] font-black bg-black/30 border border-rose-500/20 text-rose-400 hover:bg-rose-500/20 transition">🗑 Remove</button></div>
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
                <div className="rounded-2xl border border-[#FDB51F]/20 bg-[#FDB51F]/8 px-4 py-3">
                  <p className="text-[11px] font-bold leading-5 text-[#FFE29A]">📌 القواعد تُملأ تلقائياً من قالب ثابت قابل للتعديل، وتشمل مدة الفيديو (3 دقائق) والحد الأقصى للحجم (50MB).</p>
                  <button type="button" onClick={()=>{const t=rulesForPreset(selPreset&&selPreset!=="custom"?Number(selPreset):undefined,nAr||"المسابقة",nEn||"Competition");setRAr(t.ar);setREn(t.en)}} className="mt-2 rounded-full border border-white/15 bg-white/5 px-3 py-1.5 text-[11px] font-black text-white/70 hover:bg-white/10">↺ إعادة القالب القياسي • Reset template</button>
                </div>
                <div><FL>Rules (Arabic) • قواعد المسابقة</FL><TA rows={8} value={rAr} onChange={e=>setRAr(e.target.value)} dir="rtl"/></div>
                <div><FL>Rules (English)</FL><TA rows={8} value={rEn} onChange={e=>setREn(e.target.value)}/></div>
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

        {editClassesFor&&<TeacherClassesModal teacher={editClassesFor} onSave={saveTeacherClasses} onClose={()=>setEditClassesFor(null)}/>}


        {tab==="security"&&(
          <div className="grid gap-6 lg:grid-cols-2">
            <div className="rounded-[2rem] border border-white/10 bg-white/5 p-6 backdrop-blur-sm">
              <h2 className="text-xl font-black mb-1 text-white">🔑 Change Admin Password</h2>
              <p className="text-sm text-white/45 mb-5">تغيير كلمة مرور الإدارة — استخدم كلمة المرور الحالية أو رمز الاستعادة.</p>
              <div className="space-y-3">
                <div><FL>Current password OR recovery code • كلمة المرور الحالية أو رمز الاستعادة</FL><Inp type="password" value={pwOld} onChange={e=>{setPwOld(e.target.value);setPwErr("");setPwMsg("")}} placeholder="•••••••••"/></div>
                <div><FL>New password • كلمة المرور الجديدة</FL><Inp type="password" value={pwNew} onChange={e=>{setPwNew(e.target.value);setPwErr("");setPwMsg("")}} placeholder="6+ characters"/></div>
                <div><FL>Confirm new password • تأكيد كلمة المرور</FL><Inp type="password" value={pwConfirm} onChange={e=>{setPwConfirm(e.target.value);setPwErr("");setPwMsg("")}} placeholder="Repeat the new password"/></div>
                {pwErr&&<div className="rounded-2xl border border-rose-400/25 bg-rose-500/10 px-4 py-3 text-sm font-bold text-rose-200">{pwErr}</div>}
                {pwMsg&&<div className="rounded-2xl border border-emerald-400/25 bg-emerald-500/10 px-4 py-3 text-sm font-bold text-emerald-200">{pwMsg}</div>}
                <button type="button" onClick={()=>{
                  setPwErr("");setPwMsg("");
                  const old=pwOld.trim();
                  if(!(old===adminPw || old===ADMIN_RECOVERY_CODE)){setPwErr("كلمة المرور الحالية أو رمز الاستعادة غير صحيح • Wrong password or recovery code");return}
                  if(pwNew.trim().length<6){setPwErr("كلمة المرور الجديدة قصيرة (6 أحرف على الأقل) • New password too short");return}
                  if(pwNew.trim()!==pwConfirm.trim()){setPwErr("تأكيد كلمة المرور غير مطابق • Passwords do not match");return}
                  setAdminPw(pwNew.trim());
                  setPwOld("");setPwNew("");setPwConfirm("");
                  setPwMsg("✅ تم تغيير كلمة المرور وحفظها سحابياً • Password updated and synced");
                }} className="w-full rounded-2xl bg-[#FDB51F] py-4 text-sm font-black text-slate-950 hover:bg-[#F7C84A] shadow-lg">Save New Password ✓</button>
              </div>
            </div>
            <div className="space-y-6">
              <div className="rounded-[2rem] border border-white/10 bg-white/5 p-6 backdrop-blur-sm">
                <h3 className="text-lg font-black mb-1 text-white">🆘 Forgot the password?</h3>
                <p className="text-sm text-white/45 mb-4">نسيت كلمة المرور؟ استخدم رمز الاستعادة أو أرسله إلى بريد الإدارة.</p>
                <div className="rounded-2xl border border-[#FDB51F]/25 bg-[#FDB51F]/10 p-4">
                  <p className="text-[11px] font-black uppercase tracking-widest text-[#FFE29A]">Recovery code • رمز الاستعادة</p>
                  <div className="mt-2 flex items-center gap-2">
                    <code className="flex-1 rounded-xl bg-black/30 px-3 py-2 text-sm font-black tracking-widest text-white">{showCode?ADMIN_RECOVERY_CODE:"••••••••••••"}</code>
                    <button onClick={()=>setShowCode(v=>!v)} className="rounded-xl border border-white/15 bg-white/5 px-3 py-2 text-[11px] font-black text-white/70 hover:bg-white/10">{showCode?"Hide":"Show"}</button>
                    <button onClick={()=>{try{navigator.clipboard.writeText(ADMIN_RECOVERY_CODE)}catch{}}} className="rounded-xl border border-white/15 bg-white/5 px-3 py-2 text-[11px] font-black text-white/70 hover:bg-white/10">📋 Copy</button>
                  </div>
                </div>
                <div className="mt-4 rounded-2xl border border-white/10 bg-black/20 p-4">
                  <p className="text-[11px] font-bold text-white/50">Admin emails • بريد الإدارة</p>
                  <p className="text-sm font-black text-white break-all" dir="ltr">{ADMIN_EMAIL}</p>
                  <p className="mb-3 text-sm font-black text-white/70 break-all" dir="ltr">{ADMIN_EMAIL_ALT}</p>
                  <button onClick={()=>{
                    const body=[`GWA Champions — Admin access details`,`بيانات دخول الإدارة`,``,`Recovery code: ${ADMIN_RECOVERY_CODE}`,`Current password: ${adminPw}`,``,`Portal: https://gwa-champions-platform.vercel.app`].join("\n");
                    window.location.href=`mailto:${ADMIN_EMAIL},${ADMIN_EMAIL_ALT}?subject=${encodeURIComponent("GWA Champions — Admin password / recovery code")}&body=${encodeURIComponent(body)}`;
                  }} className="w-full rounded-2xl bg-[#1697E8] py-3.5 text-sm font-black text-white hover:bg-sky-400">📧 Send password &amp; code to {ADMIN_EMAIL}</button>
                  <p className="mt-2 text-[10px] leading-5 text-white/30">يفتح برنامج البريد لديك برسالة جاهزة إلى بريد الإدارة تحتوي كلمة المرور الحالية ورمز الاستعادة — أرسلها لتصل إليك عند نسيانها.</p>
                </div>
              </div>
              <div className="rounded-[2rem] border border-white/10 bg-white/5 p-6 backdrop-blur-sm">
                <h3 className="text-base font-black mb-2 text-white">ℹ️ Notes • ملاحظات</h3>
                <ul className="space-y-1.5 text-[12px] leading-6 text-white/50">
                  <li>• كلمة مرور الإدارة محفوظة سحابياً، فتعمل على كل الأجهزة.</li>
                  <li>• رمز الاستعادة ثابت ويُستخدم فقط عند نسيان الكلمة القديمة.</li>
                  <li>• بعد التغيير استخدم الكلمة الجديدة للدخول إلى بوابة الإدارة.</li>
                </ul>
              </div>
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
  const[editWId,setEditWId]=useState<string|null>(null);
  
  return(
    <div className="relative min-h-screen bg-[#0d1230] text-white"><Bg/>
      <header className="sticky top-0 z-20 border-b border-white/10 bg-[#0f1535]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center"><div><div className="text-[10px] font-bold uppercase tracking-[0.28em] text-sky-300">Teacher Account</div><h1 className="text-xl font-black">{teacher.fullName}</h1></div><button onClick={onBack} className="rounded-full border border-white/10 bg-white/5 px-5 py-2.5 text-sm font-bold text-white hover:bg-white/10">Back</button></header>
      <main className="max-w-4xl mx-auto px-6 py-10"><div className="grid gap-4 md:grid-cols-3 mb-10"><Stat l="Students" v={myStudents.length} s="طلابي"/><Stat l="Works" v={myWorks.length} s="الأعمال"/><Stat l="Criteria" v="10" s="عناصر"/></div>
        <div className="grid gap-8 md:grid-cols-2 xl:grid-cols-3">{myWorks.map(w=>{const stu=myStudents.find(s=>s.id===w.studentId);const co=comps.find(c=>c.id===w.competitionId);const ev=evals.find(e=>e.workId===w.id);if(!stu)return null;return(
          <article key={w.id} className="rounded-3xl border border-white/10 bg-white/5 overflow-hidden"><SecureVideo path={w.videoUrl} poster={w.thumb} controls className="w-full aspect-video object-cover"/><div className="p-5"><div className="text-[10px] font-black uppercase text-sky-300">{co?.nameEn}</div><h3 className="text-base font-black text-white mt-1">{w.title}</h3><p className="text-sm text-[#FDB51F] mt-1">👤 {stu.fullName}</p><div className="mt-4"><div className="w-full rounded-2xl bg-white/5 border border-white/10 px-4 py-3 text-sm font-bold text-white/50 text-center">Use Evaluation Code on Works ⬆</div></div></div></article>);})}</div></main>
    </div>
  );
}

/* ── TEACHER LOGIN QUICK MODAL ── */
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
          <div className="rounded-2xl border border-white/10 bg-white/5 p-5 space-y-5"><div className="border-l-4 border-[#1697E8] pl-4"><p className="text-sm font-black text-[#1697E8] mb-2">📌 Declaration in English</p><p className="text-sm leading-7 text-white/70">I, <b className="text-white">{student.fullName}</b>, confirm my voluntary participation in <b className="text-[#1697E8]">"{compEn}"</b>.</p><ul className="mt-2 space-y-1 text-sm leading-7 text-white/55"><li>✅ My video will be evaluated fairly by a panel of teachers.</li><li>⚠️ My video stays locked ("Waiting Teachers Approval") until a teacher approves it — then it appears on the Champions page.</li>
                <li>⏱ Max duration 3 minutes • 📦 Max file size 50 MB.</li><li>✅ The submitted work is my own original creation.</li></ul></div>
            <div className="border-t border-white/10"/><div dir="rtl" className="border-r-4 border-[#FDB51F] pr-4"><p className="text-sm font-black text-[#FDB51F] mb-2">📌 إقرار بالعربية</p><p className="text-sm leading-8 text-white/80">أنا <b className="text-white">{student.fullName}</b>، أُقر بمشاركتي الطوعية في مسابقة <b className="text-[#FDB51F]">«{compAr}»</b>.</p><ul className="mt-2 space-y-1 text-sm leading-8 text-white/70"><li>✅ أعلم أن الفيديو سيُعرض للتقييم العادل من لجنة المعلمين.</li><li>⚠️ أعلم أن الفيديو يبقى مغلقاً بعبارة «بانتظار اعتماد المعلم» حتى يعتمده أحد المعلمين ثم يظهر للجميع.</li>
                <li>⏱ مدة الفيديو ٣ دقائق كحد أقصى • 📦 الحد الأقصى للحجم 50 ميجابايت.</li><li>✅ أؤكد أن العمل المقدم هو عملي الأصلي.</li></ul></div></div>
          <label className="mt-5 flex cursor-pointer items-start gap-4 rounded-2xl border border-[#FDB51F]/20 bg-[#FDB51F]/8 p-4 hover:bg-[#FDB51F]/12"><input type="checkbox" checked={ok} onChange={e=>setOk(e.target.checked)} className="mt-1 h-5 w-5 accent-[#FDB51F]"/><div><p className="text-sm font-black text-[#FFE29A]">I confirm and agree to all terms above.</p><p className="text-sm text-white/40 mt-0.5" dir="rtl">أؤكد رغبتي في المشاركة وأوافق على جميع البنود.</p></div></label>
          <div className="mt-5 flex gap-3"><button onClick={onNo} className="flex-1 rounded-2xl border border-white/10 bg-white/5 px-5 py-4 text-sm font-black text-white/70 hover:bg-white/10">Decline</button><button disabled={!ok} onClick={onOk} className={cn("flex-[2] rounded-2xl px-5 py-4 text-sm font-black shadow-lg",ok?"bg-[#FDB51F] text-slate-950 hover:bg-[#F7C84A]":"cursor-not-allowed bg-white/10 text-white/30")}>✅ Confirm & Upload</button></div>
        </div>
      </div>
    </div>
  );
}

function StudentPage({student,comps,works,evals,certs,selCompId,setWorks,onBack}:{student:User;comps:Comp[];works:Work[];evals:Ev[];certs:Cert[];selCompId:string|null;setWorks:React.Dispatch<React.SetStateAction<Work[]>>;onBack:()=>void}){
  // Update available competitions to allow ALL active comps if student grade doesn't exactly match
  // or loosely match "5" vs "Grade 5"
  const avail=useMemo(()=>{
    const active = comps.filter(c=>c.started&&!c.archived);
    // Teachers signing in through the Student portal may upload to any active competition
    if(student.role==="teacher") return active;
    const hasExact = active.filter(c=>c.grades.includes(student.grade));
    if (hasExact.length > 0) return hasExact;
    
    // Loose matching ("5" matches "Grade 5")
    const looseGrade = student.grade.replace(/[^0-9]/g, "");
    const hasLoose = active.filter(c=>c.grades.some(g=>g.replace(/[^0-9]/g, "") === looseGrade));
    if (hasLoose.length > 0) return hasLoose;
    
    // Fallback: show all if no grade matches
    return active;
  },[comps,student.grade]);
  const myWorks=useMemo(()=>works.filter(w=>w.studentId===student.id),[student.id,works]);
  const myCerts=useMemo(()=>certs.filter(c=>c.studentId===student.id),[certs,student.id]);
  const[sTab,setSTab]=useState<"participate"|"submissions"|"certs">("participate");
  const[step,setStep]=useState<"idle"|"consent"|"form">("idle");
  const[cid,setCid]=useState(selCompId??avail[0]?.id??"");
  const[title,setTitle]=useState(""); const[desc,setDesc]=useState(""); const[vidFile,setVidFile]=useState<File|null>(null); const[vidUrl,setVidUrl]=useState("");
  const[prog,setProg]=useState(0); const[upl,setUpl]=useState(false); const[msg,setMsg]=useState("");
  useEffect(()=>{if(selCompId)setCid(selCompId)},[selCompId]);
  function onVid(e:React.ChangeEvent<HTMLInputElement>){
    const f=e.target.files?.[0]; if(!f) return;
    if(f.size > MAX_VIDEO_MB*1024*1024){
      setMsg(`❌ حجم الفيديو ${(f.size/1048576).toFixed(1)}MB أكبر من الحد الأقصى ${MAX_VIDEO_MB}MB • File too large`);
      setVidFile(null); setVidUrl(""); return;
    }
    const probe=document.createElement("video");
    probe.preload="metadata";
    probe.onloadedmetadata=()=>{
      const dur=probe.duration||0;
      if(dur > MAX_VIDEO_SECONDS+3){
        setMsg(`❌ مدة الفيديو ${Math.round(dur)} ثانية أكبر من ${MAX_VIDEO_SECONDS/60} دقائق • Video longer than 3 minutes`);
        setVidFile(null); setVidUrl("");
      }else{
        setMsg(""); setVidFile(f); setVidUrl(URL.createObjectURL(f));
      }
    };
    probe.onerror=()=>{ setMsg(""); setVidFile(f); setVidUrl(URL.createObjectURL(f)); };
    probe.src=URL.createObjectURL(f);
  }
  async function submit(e:React.FormEvent){
    e.preventDefault();
    if(!cid || !title.trim() || !vidFile){setMsg("All fields required"); return}
    if(vidFile.size > MAX_VIDEO_MB*1024*1024){setMsg(`❌ الحد الأقصى لحجم الفيديو ${MAX_VIDEO_MB}MB • File too large`); return}
    
    setUpl(true);
    setProg(0);

    try {
      // Upload the video to cloud storage (Netlify Blobs)
      const videoUrl = await uploadVideo(vidFile, setProg);

      // Save the submission record to the cloud database
      const work: Work = {
        id: mid("w"),
        competitionId: cid,
        studentId: student.id,
        title: title.trim(),
        desc: desc.trim(),
        videoUrl,
        thumb: "",
        at: new Date().toISOString(),
        status: "pending" as WorkStatus
      };
      // Supabase is the source of truth; the legacy endpoint is best-effort only.
      try {
        await fetch("/api/works", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(work)
        });
      } catch { /* ignore legacy endpoint failures */ }

      setWorks(p => [work, ...p]);

      setMsg("🎉 تم الإرسال! الفيديو الآن بانتظار اعتماد المعلم — Waiting Teachers Approval 1st");
      setTitle(""); setDesc(""); setVidFile(null); setVidUrl(""); setStep("idle");
    } catch (err: any) {
      setMsg("❌ Upload Error: " + err.message);
    } finally {
      setUpl(false);
    }
  }
  const selComp=avail.find(c=>c.id===cid);
  
  function printCert(cert:Cert){
    const win=window.open("","_blank","width=900,height=700");
    if(!win)return;
    win.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8"/><title>Certificate</title>
    <style>
      *{box-sizing:border-box;margin:0;padding:0}body{display:flex;align-items:center;justify-content:center;min-height:100vh;background:#f0f0f0;font-family:"Segoe UI",Arial,sans-serif;padding:20px}
      .cert{width:800px;background:#fff;border:4px solid #1F2E76;border-radius:24px;padding:48px;text-align:center;position:relative;overflow:hidden}
      .cert::before{content:"";position:absolute;top:0;left:0;right:0;height:8px;background:linear-gradient(90deg,#FDB51F,#1697E8,#10B09B,#E21455)}
      .logo{width:80px;margin:0 auto 16px}
      .school{font-size:14px;color:#1F2E76;font-weight:700;letter-spacing:3px;text-transform:uppercase}
      .platform{font-size:12px;color:#FDB51F;font-weight:700;margin-top:4px}
      h1{font-size:36px;color:#1F2E76;margin:24px 0 8px;font-weight:900}
      h2{font-size:20px;color:#555;font-weight:600}
      .name{font-size:32px;color:#1F2E76;font-weight:900;margin:24px 0;padding:12px 0;border-top:2px solid #eee;border-bottom:2px solid #eee}
      .details{font-size:16px;color:#555;line-height:2}
      .rank{display:inline-block;background:linear-gradient(135deg,#FDB51F,#E21455);color:#fff;padding:8px 24px;border-radius:999px;font-size:18px;font-weight:900;margin:20px 0}
      .score{font-size:48px;font-weight:900;color:#1F2E76;margin:16px 0}
      .footer{font-size:11px;color:#aaa;margin-top:32px}
      .tools{text-align:center;margin-bottom:16px}
      .tools button{padding:10px 20px;border-radius:999px;border:none;background:#FDB51F;font-weight:700;cursor:pointer;margin:0 4px}
      @media print{.tools{display:none}body{background:#fff;padding:0}}
    </style></head><body>
    <div>
      <div class="tools"><button onclick="window.print()">🖨️ Print</button><button onclick="window.close()">Close</button></div>
      <div class="cert">
        <img src="/gems-logo.png" class="logo" alt="GEMS"/>
        <div class="school">GEMS World Academy Dubai</div>
        <div class="platform">GWA Arabic & Islamic Champions<br/>أبطال اللغة العربية والتربية الإسلامية</div>
        <h1>Certificate of Excellence</h1>
        <h2>شهادة تقدير وتميز</h2>
        <div class="name">${cert.compName.replace(/</g,"&lt;")}</div>
        <div class="details">
          Awarded to: <strong>${student.fullName.replace(/</g,"&lt;")}</strong><br/>
          ${cert.grade} • Section ${cert.section}<br/>
        </div>
        <div class="rank">${MEDALS[cert.rank-1]||"#"+cert.rank} Rank #${cert.rank}</div>
        <div class="score">${cert.score}/100</div>
        <div class="footer">
          Issued: ${cert.issuedAt} • GWA Champions Platform • أكاديمية جيمس العالمية
        </div>
      </div>
    </div>
    </body></html>`);
    win.document.close();
  }

  return(
    <div className="relative min-h-screen bg-[#0d1230] text-white"><Bg/>
      {step==="consent"&&selComp&&<ConsentModal student={student} compAr={selComp.nameAr} compEn={selComp.nameEn} onOk={()=>setStep("form")} onNo={()=>setStep("idle")}/>}
      <header className="sticky top-0 z-20 border-b border-white/10 bg-[#0f1535]/80 backdrop-blur-xl px-6 py-4 flex justify-between items-center"><div><div className="text-[10px] font-bold uppercase tracking-[0.28em] text-[#FDB51F]">Student Portal</div>{student.role==="teacher"&&<div className="mt-1 inline-block rounded-full bg-[#FDB51F]/15 border border-[#FDB51F]/30 px-3 py-1 text-[10px] font-black text-[#FFE29A]">👩‍🏫 Teacher upload mode — uploads need another teacher's approval</div>}<h1 className="text-xl font-black">{student.fullName}</h1><p className="text-sm text-white/50">{student.grade} • {student.section}</p></div><button onClick={onBack} className="rounded-full border border-white/10 bg-white/5 px-5 py-2.5 text-sm font-bold text-white hover:bg-white/10">Back</button></header>
      <main className="max-w-6xl mx-auto px-6 py-8">
        <div className="flex flex-wrap gap-2 mb-8 bg-white/5 border border-white/10 rounded-2xl p-1.5">
          <button onClick={()=>setSTab("participate")} className={cn("rounded-xl px-5 py-3 text-sm font-black",sTab==="participate"?"bg-[#FDB51F] text-slate-950":"text-white/50")}>🎓 Participate</button>
          <button onClick={()=>setSTab("submissions")} className={cn("rounded-xl px-5 py-3 text-sm font-black",sTab==="submissions"?"bg-[#1697E8] text-white":"text-white/50")}>📁 My Works ({myWorks.length})</button>
          <button onClick={()=>setSTab("certs")} className={cn("rounded-xl px-5 py-3 text-sm font-black",sTab==="certs"?"bg-emerald-500 text-white":"text-white/50")}>📜 Certificates ({myCerts.length})</button>
        </div>
        {sTab==="participate"&&(
          <div className="max-w-xl"><div className="rounded-[2rem] border border-white/10 bg-white/5 p-6 backdrop-blur-sm">
            {step==="idle"&&(<><h2 className="text-xl font-black mb-4">Participate</h2><div className="space-y-4"><div><FL>Select Competition</FL><Sel value={cid} onChange={e=>{setCid(e.target.value);setMsg("")}}><option value="">{avail.length===0?"No active competitions for your grade":"— Select —"}</option>{avail.map(c=><option key={c.id} value={c.id}>{c.nameEn}</option>)}</Sel></div>{msg&&<div className="p-3 bg-emerald-500/10 border border-emerald-400/20 text-emerald-200 rounded-xl text-sm leading-6">{msg}</div>}
              {selComp&&(<div className="rounded-2xl border border-white/10 bg-white/5 p-4">
                <p className="text-sm font-black text-[#FDB51F] mb-2">📋 قواعد المسابقة • Competition rules</p>
                <pre className="whitespace-pre-wrap break-words text-xs leading-6 text-white/60 font-sans" dir="rtl">{selComp.rulesAr||"—"}</pre>
                <div className="mt-3 flex flex-wrap gap-2">
                  <Pill t="gold">⏱ Max 3 minutes</Pill>
                  <Pill t="blue">📦 Max {MAX_VIDEO_MB} MB</Pill>
                </div>
              </div>)}
              <button disabled={!cid} onClick={()=>{setMsg("");setStep("consent")}} className={cn("w-full rounded-2xl px-5 py-4 text-sm font-black",cid?"bg-[#FDB51F] text-slate-950 hover:bg-[#F7C84A] shadow-lg":"cursor-not-allowed bg-white/10 text-white/30")}>📋 Read Consent & Participate</button></div></>)}
            {step==="form"&&(<form onSubmit={submit} className="space-y-5"><div className="flex items-center gap-3 mb-2"><button type="button" onClick={()=>setStep("idle")} className="text-white/50 text-xs">← Back</button><p className="text-xs text-emerald-300 font-bold">✅ Consent Accepted</p></div><div><FL>Title</FL><Inp value={title} onChange={e=>setTitle(e.target.value)} placeholder="Work title"/></div><div><FL>📹 Upload Video</FL><label className={cn("flex flex-col items-center gap-3 rounded-2xl border-2 border-dashed p-6 cursor-pointer transition",vidFile?"border-emerald-400/40 bg-emerald-500/5":"border-white/10 bg-white/5 hover:border-white/25")}><input type="file" accept="video/mp4,video/quicktime,video/*" onChange={onVid} className="hidden"/>{vidFile?(<><span className="text-3xl">🎬</span><p className="font-black text-emerald-300">{vidFile.name}</p><video src={vidUrl} controls className="w-full max-h-28 rounded-xl mt-2"/></>):(<><span className="text-3xl">📹</span><p className="text-sm font-black text-white">Click to select</p><p className="text-xs text-white/40">MP4, MOV, AVI • ⏱ Max 3 minutes • 📦 Max 50 MB</p></>)}</label></div>{upl&&<div className="rounded-2xl border border-sky-400/20 bg-sky-500/10 p-5 text-center animate-pulse"><div className="flex items-center justify-center gap-3 text-sky-200"><span className="text-xl">📹</span><span className="text-sm font-black uppercase tracking-widest">Uploading Video...</span></div><p className="text-[10px] text-white/40 mt-1 uppercase font-bold tracking-tighter">Please do not close the browser</p></div>}<button type="submit" disabled={upl} className={cn("w-full rounded-2xl px-5 py-4 text-sm font-black shadow-lg",upl?"cursor-not-allowed bg-white/10 text-white/30":"bg-[#FDB51F] text-slate-950 hover:bg-[#F7C84A]")}>{upl?`${prog}%...`:"🚀 Submit Work"}</button></form>)}
          </div></div>
        )}
        {sTab==="submissions"&&(<div className="grid gap-6 md:grid-cols-2 xl:grid-cols-3">{myWorks.length===0&&<div className="md:col-span-3 rounded-3xl border border-white/10 bg-white/5 p-12 text-center text-white/40">No submissions yet</div>}{myWorks.map(w=>{const ev=evals.find(e=>e.workId===w.id);const co=comps.find(c=>c.id===w.competitionId);return(<article key={w.id} className="rounded-2xl border border-white/10 bg-white/5 overflow-hidden"><SecureVideo path={w.videoUrl} poster={w.thumb} controls className="w-full aspect-video object-cover"/><div className="p-4"><div className="text-[10px] font-black uppercase text-sky-300">{co?.nameEn}</div><h3 className="text-base font-black text-white mt-1">{w.title}</h3><div className="mt-3 flex flex-wrap gap-2"><StatusPill w={w}/>{ev?<Pill t="green">✅ Score: {tot(ev.scores)}/100</Pill>:null}</div>{!isApproved(w)&&<p className="mt-2 text-[11px] leading-5 text-white/40" dir="rtl">هذه معاينة خاصة لك فقط. لن يظهر الفيديو للجميع حتى يعتمده أحد المعلمين.</p>}</div></article>);})}</div>)}
        {sTab==="certs"&&(<div className="space-y-6">{myCerts.length===0&&<div className="rounded-3xl border border-white/10 bg-white/5 p-12 text-center"><p className="text-4xl mb-3">📜</p><p className="text-white/40">No certificates yet.</p></div>}{myCerts.map(cert=><article key={cert.id} className="rounded-3xl border border-[#FDB51F]/30 bg-gradient-to-r from-[#FDB51F]/10 to-[#0d1230] p-6 flex items-center gap-5 justify-between flex-wrap"><div className="flex items-center gap-4"><div className="text-4xl">{MEDALS[cert.rank-1]||"🏅"}</div><div><h3 className="text-xl font-black text-white">{cert.compName}</h3><div className="flex gap-2 mt-2"><Pill t="gold">Rank #{cert.rank}</Pill><Pill t="green">Score: {cert.score}/100</Pill><Pill t="blue">{cert.grade}</Pill></div></div></div><button onClick={()=>printCert(cert)} className="bg-[#FDB51F] text-slate-950 px-5 py-3 rounded-2xl font-black">Download PDF</button></article>)}</div>)}
      </main>
    </div>
  );
}