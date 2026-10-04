import type { CompPreset } from "../types";

/** localStorage keys — will become database table names in Phase 3 */
export const STORAGE_KEYS = {
  users: "gwa-u-v8",
  comps: "gwa-c-v8",
  works: "gwa-w-v8",
  evals: "gwa-e-v8",
  likes: "gwa-l-v8",
  certs: "gwa-cert-v8",
  archive: "gwa-arch-v8",
  adminPw: "gwa-pw-v8",
};

export const GRADES = [
  "KG 1", "KG 2",
  "Grade 1", "Grade 2", "Grade 3", "Grade 4", "Grade 5", "Grade 6",
  "Grade 7", "Grade 8", "Grade 9", "Grade 10", "Grade 11", "Grade 12",
];

export const SECTION_OPTIONS = ["A", "B", "C", "D", "X", "Y"];

export const DEFAULT_CRITERIA = [
  "Creativity • الإبداع", "Clarity • الوضوح", "Content • المحتوى",
  "Language • اللغة", "Confidence • الثقة", "Design • التصميم",
  "Compliance • الالتزام", "Organization • التنظيم",
  "Technology • التقنية", "Final Impact • الإخراج",
];

export const MEDALS = ["🥇", "🥈", "🥉", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"];

export const COMP_PRESETS: CompPreset[] = [
  { nameEn:"Quran Recitation Competition", nameAr:"مسابقة تلاوة القرآن الكريم", criteria:["Tajweed Rules • أحكام التجويد","Letter Pronunciation • مخارج الحروف","Voice Beauty • جمال الصوت","Pausing & Starting • الوقف والابتداء","Emotional Expression • تصوير المعنى","Breath Control • التحكم بالنفس","Memorization Accuracy • دقة الحفظ","Rhythm & Flow • الإيقاع والانسيابية","Confidence • الثقة بالنفس","Overall Impression • الانطباع العام"] },
  { nameEn:"Arabic Public Speaking", nameAr:"مسابقة التحدث بالعربية الفصحى", criteria:["Pronunciation • سلامة النطق","Grammar Accuracy • السلامة النحوية","Vocabulary Richness • ثراء المفردات","Fluency • الطلاقة","Body Language • لغة الجسد","Eye Contact • التواصل البصري","Content Depth • عمق المحتوى","Audience Engagement • تفاعل الجمهور","Time Management • إدارة الوقت","Persuasion • القدرة على الإقناع"] },
  { nameEn:"Arabic Creative Writing", nameAr:"مسابقة الكتابة الإبداعية", criteria:["Original Idea • أصالة الفكرة","Language Style • أسلوب اللغة","Grammar & Spelling • النحو والإملاء","Story Structure • بنية القصة","Character Development • تطور الشخصيات","Imagery & Description • التصوير والوصف","Emotional Impact • التأثير العاطفي","Vocabulary Range • تنوع المفردات","Creativity • الإبداع","Conclusion • قوة الخاتمة"] },
  { nameEn:"Arabic Poetry Recitation", nameAr:"مسابقة إلقاء الشعر العربي", criteria:["Pronunciation • سلامة النطق","Rhythm & Meter • الوزن والإيقاع","Emotional Delivery • الأداء العاطفي","Voice Modulation • تنويع الصوت","Memorization • الحفظ","Stage Presence • الحضور المسرحي","Understanding of Meaning • فهم المعنى","Audience Connection • التواصل مع الجمهور","Gestures • الإيماءات","Overall Performance • الأداء الكلي"] },
  { nameEn:"Arabic Reading Competition", nameAr:"مسابقة القراءة العربية", criteria:["Reading Fluency • طلاقة القراءة","Comprehension • الفهم والاستيعاب","Critical Analysis • التحليل النقدي","Summary Skills • مهارات التلخيص","Vocabulary Usage • استخدام المفردات","Presentation Quality • جودة العرض","Book Selection • اختيار الكتاب","Personal Reflection • التأمل الشخصي","Research Depth • عمق البحث","Creative Expression • التعبير الإبداعي"] },
  { nameEn:"Arabic Calligraphy", nameAr:"مسابقة الخط العربي", criteria:["Letter Form • شكل الحروف","Consistency • التناسق","Spacing • المسافات","Line Quality • جودة الخط","Artistic Composition • التكوين الفني","Creativity • الإبداع","Color Usage • استخدام الألوان","Neatness • النظافة","Adherence to Rules • الالتزام بالقواعد","Overall Aesthetics • الجمال العام"] },
  { nameEn:"Islamic Knowledge Quiz", nameAr:"مسابقة المعارف الإسلامية", criteria:["Quran Knowledge • معرفة القرآن","Hadith Knowledge • معرفة الحديث","Fiqh Understanding • فهم الفقه","Seerah Knowledge • معرفة السيرة","Islamic History • التاريخ الإسلامي","Speed of Answer • سرعة الإجابة","Accuracy • الدقة","Explanation Quality • جودة التفسير","Confidence • الثقة","Sportsmanship • الروح الرياضية"] },
  { nameEn:"Arabic Spelling Bee", nameAr:"مسابقة الإملاء العربي", criteria:["Spelling Accuracy • دقة الإملاء","Dictation Speed • سرعة الإملاء","Hamza Rules • قواعد الهمزة","Ta Marbuta/Ha • التاء المربوطة والهاء","Alef Rules • قواعد الألف","Punctuation • علامات الترقيم","Word Derivation • اشتقاق الكلمات","Handwriting • جودة الخط","Listening Skills • مهارات الاستماع","Overall Performance • الأداء الكلي"] },
  { nameEn:"Social Studies Project", nameAr:"مشروع الدراسات الاجتماعية", criteria:["Research Quality • جودة البحث","Topic Relevance • ملاءمة الموضوع","Data Presentation • عرض البيانات","Critical Thinking • التفكير النقدي","Creativity • الإبداع","Visual Design • التصميم البصري","Historical Accuracy • الدقة التاريخية","Oral Presentation • العرض الشفهي","Teamwork • العمل الجماعي","Conclusions • الاستنتاجات"] },
  { nameEn:"Arabic Debate Competition", nameAr:"مسابقة المناظرة بالعربية", criteria:["Argument Strength • قوة الحجة","Rebuttal Skills • مهارات الرد","Evidence Usage • استخدام الأدلة","Language Proficiency • إتقان اللغة","Persuasion • الإقناع","Logical Structure • البنية المنطقية","Delivery Style • أسلوب الإلقاء","Time Management • إدارة الوقت","Teamwork • العمل الجماعي","Overall Impact • التأثير العام"] },
];
