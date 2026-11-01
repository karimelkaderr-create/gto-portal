// GTO Assist – answers agents' questions from the Promotions & Products knowledge base.
// Runs entirely on the portal server: no external AI service, no cost. Every answer quotes the sheet.
// Steps: understand the language → find the offer(s) meant → work out what is being asked →
// pull the exact field plus related sentences from anywhere in that offer → reply with links and follow-ups.

/* ---------- text helpers ---------- */
const AR_DIAC = /[ً-ْـ]/g;
const normAr = s => s.replace(AR_DIAC, '').replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه').replace(/ؤ/g, 'و').replace(/ئ/g, 'ي');
const isArabic = s => (String(s).match(/[؀-ۿ]/g) || []).length > (String(s).match(/[a-z]/gi) || []).length;
// Arabic (and common Franco/English spellings) of product words → the English word used in the sheet
const TRANSLIT = {
  ماجيك: 'magic', ماچيك: 'magic', باس: 'pass', باص: 'pass', فلايت: 'flight', فلاي: 'flight', كلوب: 'club', كلاب: 'club', هوليداي: 'holiday', هوليدي: 'holiday',
  كارد: 'card', كارت: 'card', فاوتشر: 'voucher', فوتشر: 'voucher', فاوشر: 'voucher', قسيمه: 'voucher', سمارت: 'smart', بوش: 'posh', سيفر: 'saver', سيڤر: 'saver',
  كودوس: 'kudos', كيودوس: 'kudos', ترافل: 'travel', باوند: 'pound', باوندز: 'pound', ونتر: 'winter', وينتر: 'winter', سمر: 'summer', صيفي: 'summer', فلاش: 'flash',
  تريب: 'trip', كي: 'key', نكست: 'next', داي: 'day', توموروا: 'tomorrow', تومورو: 'tomorrow', فاست: 'fast', ارجنت: 'urgent', فاميلي: 'family', جونيور: 'junior',
  بلاك: 'black', فرايداي: 'friday', كريسماس: 'christmas', عيد: 'eid', عمره: 'umrah', اوروبا: 'europe', يوروب: 'europe', شوب: 'shop', داين: 'dine', ايميرالد: 'emerald',
  امرالد: 'emerald', توتال: 'total', سيكريت: 'secret', سيكريتس: 'secret', كونسيرج: 'concierge', ايربنب: 'airbnb', ايفون: 'iphone',
  اتصالات: 'etisalat', ايتصالات: 'etisalat', فاليو: 'valu', ماستركارد: 'mastercard', فيزا: 'visa', اتش: 'hsbc', بانك: 'bank', كوينز: 'currency',
  فندق: 'hotel', فنادق: 'hotel', اوتيل: 'hotel', اوتيلات: 'hotel', طيران: 'flight', طياره: 'flight', تذكره: 'ticket', تذاكر: 'ticket', رحله: 'trip', رحلات: 'trip',
  خصم: 'discount', خصومات: 'discount', عرض: 'offer', عروض: 'offer', محفظه: 'wallet', رصيد: 'credit', غرفه: 'room', غرف: 'room', مجاني: 'free', مجانا: 'free', ببلاش: 'free',
  عضويه: 'membership', اشتراك: 'membership', جنيه: 'egp', استرليني: 'gbp', دولار: 'usd', تعليم: 'education', مدارس: 'education', فواتير: 'invoice', فاتوره: 'invoice'
};
const STOP = new Set(('a an the and or of to in on for with by at from is are be was were it its this that these those do does did can could should would will shall may might must ' +
  'i we you he she they them his her their our my me us member members customer customers client clients agent please pls plz hi hello thanks thank about tell me give what whats ' +
  'which who how when where why if there any some all much many get got have has had also just only still yet so than then as up out into over under per via ' +
  'من في علي على الى إلى عن مع هو هي هما هم انا احنا انت ده دي دا دى ايه إيه اي أي كام هل لو يا ممكن عايز عاوز عايزه محتاج العميل عميل الكلاينت كلاينت بتاع بتاعت بتاعه ' +
  'اللي الي ازاي امتي ليه فين و او ولا بس كمان برضه طيب تمام شكرا').split(/\s+/));
// words that appear in offer names but don't identify a product
const NAME_NOISE = new Set(('t c cs tc tcs terms conditions and offer offers promo promotion gto new extended extension extend shared on off up to save enjoy get your any ' +
  'the for with in of a an every one you exclusive limited time only now worldwide is are it s st nd rd th back our yearly gift x year').split(/\s+/));
const GENERIC = new Set(('hotel room stay flight ticket refundable non nr discount free egp gbp usd deposit card voucher offer sale credit wallet booking book trip travel ' +
  'international domestic attraction theme park car rental membership worldwide price pay more extra double half one two get room bonus up').split(/\s+/));
const MONTHS = { jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3, apr: 4, april: 4, may: 5, jun: 6, june: 6, jul: 7, july: 7, aug: 8, august: 8, sep: 9, sept: 9, september: 9, oct: 10, october: 10, nov: 11, november: 11, dec: 12, december: 12,
  يناير: 1, فبراير: 2, مارس: 3, ابريل: 4, مايو: 5, يونيو: 6, يوليو: 7, اغسطس: 8, سبتمبر: 9, اكتوبر: 10, نوفمبر: 11, ديسمبر: 12 };
function stem(t) {
  if (/^\d+$/.test(t) || t.length < 4 || /[؀-ۿ]/.test(t)) return t;
  if (t.endsWith('sses')) return t.slice(0, -2); if (t.endsWith('ies') && t.length > 4) return t.slice(0, -3) + 'y';
  if (t.endsWith('s') && !t.endsWith('ss') && !t.endsWith('us')) return t.slice(0, -1); return t;
}
function tokens(s) {
  let x = normAr(String(s ?? '').toLowerCase()).replace(/[’'`]/g, '').replace(/\b(2|two)\s*(for|x|4)\s*(1|one)\b/g, ' twoforone ').replace(/\b(book|buy)\s+(one|1)(\s+room)?,?\s+get\s+(one|1)(\s+room)?(\s+free)?\b|\bb1g1\b|\bbogo\w*\b/g, ' bogo ').replace(/\bhalf.?price\b/g, ' halfprice ').replace(/\be\s*&/g, ' etisalat ').replace(/\be\s+and\b/g, ' etisalat ').replace(/&/g, ' and ')
    .replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/([a-z])(\d)/g, '$1 $2').replace(/(\d)([a-z])/g, '$1 $2');
  return x.replace(/[\u060C\u061B\u061F\u066A-\u066D\u06D4]/g, ' ').split(/[^a-z0-9\u0600-\u06ff]+/).filter(Boolean).map(t => TRANSLIT[t] || TRANSLIT[t.replace(/^(ال|بال|لل|وال|فال|ب|ل|و)/, '')] || t).flatMap(t => t.split(' ')).map(t => STOP.has(t) ? t : stem(t));
}
const lev = (a, b, max) => { if (Math.abs(a.length - b.length) > max) return max + 1; let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) { const cur = [i]; let best = i; for (let j = 1; j <= b.length; j++) { cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)); best = Math.min(best, cur[j]); } if (best > max) return max + 1; prev = cur; } return prev[b.length]; };
const close = (q, n) => q === n || (q.length === 3 && n.length === 4 && n.startsWith(q) && !/^\d/.test(q) && !['car', 'par', 'for', 'the', 'day', 'one', 'way', 'all', 'pay', 'off', 'per', 'any', 'new', 'can'].includes(q)) || (q.length >= 4 && n.length >= 4 && (lev(q, n, q.length >= 8 ? 2 : 1) <= (q.length >= 8 ? 2 : 1) || (q.length >= 5 && n.startsWith(q)) || (n.length >= 5 && q.startsWith(n))));
const clean = s => String(s ?? '').replace(/\r/g, '').replace(/[ \t]+/g, ' ').trim();
const cut = (s, n) => { s = clean(s).replace(/\s*\n\s*/g, ' '); return s.length > n ? s.slice(0, n).replace(/\s+\S*$/, '') + '…' : s; };
const sentences = text => clean(text).split(/\n+|(?:^|\s)[•▪●◦]\s*|(?<=[.!?])\s+(?=[A-Z•])/).map(x => x.replace(/^(?:[-*•▪●◦]+\s*|\d{1,2}[.)]\s+)/, '').trim()).filter(x => x.length > 2);

/* ---------- field roles ---------- */
const FIELDS = [
  ['benefit', /^benefit|discount rate/i], ['offer', /offer period/i], ['booking', /booking period/i], ['expiry', /expiry/i], ['restr', /booking restrictions/i],
  ['min', /min purchase/i], ['others', /booking for others/i], ['refund', /^refundable\s*$|amount is refundable/i], ['quote', /quotation/i], ['tc', /terms/i],
  ['buy', /buying availability/i], ['rollout', /rollout/i],
  ['c0', /^refundable hotels/i], ['c1', /^nr hotels/i], ['c2', /^international flights/i], ['c3', /^domestic flights/i], ['c4', /^attraction/i], ['c5', /^theme parks/i]
];
const COVER = [
  { k: 'c0', en: 'Refundable hotels', ar: 'الفنادق القابلة للاسترداد' }, { k: 'c1', en: 'Non-refundable hotels', ar: 'الفنادق غير القابلة للاسترداد' },
  { k: 'c2', en: 'International flights', ar: 'الطيران الدولي' }, { k: 'c3', en: 'Domestic flights', ar: 'الطيران الداخلي' },
  { k: 'c4', en: 'Attractions, activities, transfers, trains & ferries', ar: 'الأنشطة والمزارات والمواصلات والقطارات والعبارات' }, { k: 'c5', en: 'Theme parks & car rental', ar: 'الملاهي وتأجير السيارات' }
];
function yn(v) {
  const s = clean(v).toLowerCase(); if (!s) return null;
  if (/^(x|no|n|✗|×|-|_)\b|^(not applicable|n\/a)\b/.test(s) || s === 'x') return false;
  if (/^not sure|^unclear|^unknown/.test(s)) return 'unsure';
  if (/^(yes|y|✓|✔|ok|included|all)\b|^\(?yes/.test(s)) return true;
  if (/^[-_—.]+$/.test(s)) return null;
  return 'partial';
}

/* ---------- question types ---------- */
// en: regex on lower-case text; ar: regex on normalised Arabic; words: what to look for inside the offer's text
const INTENTS = [
  { k: 'compare', en: /\b(vs\.?|versus|compare|comparison|difference|differ)\b/, ar: /(الفرق|قارن|مقارنه|ولا احسن|احسن من)/ },
  { k: 'min', field: 'min', en: /\b(min(imum)?|at least|lowest amount|smallest|least amount|starting (amount|from)|how much (do|does|should|must|can|to) .{0,25}(pay|deposit|buy|spend|top.?up|load|put)|minimum (deposit|purchase|spend|top.?up))\b/, ar: /(اقل|ادني|الحد الادني|حد ادني|يدفع كام|يحط كام|يشحن كام|يودع كام|اقل حاجه)/, words: /minimum|\bmin\.?\s|at least|starting from|from (gbp|egp|£)|minimum deposit/i },
  { k: 'max', en: /\b(max(imum)?|up to how much|highest amount|cap|ceiling|how much can (they|he|she|i|we|the member) (deposit|load|put|spend))\b/, ar: /(اقصي|الحد الاقصي|اكتر مبلغ|اعلي مبلغ|اكبر مبلغ)/, words: /maximum|max\.?\s|up to|cap|limit/i },
  { k: 'expiry', field: 'expiry', en: /\b(expir\w*|validity|valid (until|till)|how long is it valid|how long .{0,30}(valid|last|keep|use)|lapse|forfeit|stay in (the|their|his|her) wallet|use it until|when does (it|the (deposit|credit|balance|wallet|voucher|pass|card)) (end|expire))\b/, ar: /(صلاحي|ينتهي|بتنتهي|تنتهي|مده الاستخدام|يفضل لحد|بيخلص امتي|هتنتهي|اكسباير)/, words: /expir|valid|lapse|forfeit|remain available|no expiry|does not expire/i },
  { k: 'offer', field: 'offer', en: /\b(offer (end|period|close|valid|deadline)|when does (the )?(offer|sale|promo\w*|campaign) (end|finish|close|stop)|last (day|date|chance) to (buy|purchase|join)|until when can .{0,25}(buy|purchase|join)|deadline|sale (end|period)|end date|sales? window|purchase (period|window))\b/, ar: /(العرض (ينتهي|بينتهي|هيخلص|لحد امتي|شغال لحد)|اخر يوم|اخر ميعاد|امتي العرض|مده العرض|العرض لحد)/, words: /offer ends?|valid until|ends? at|until \d|deadline|while stock|quota/i },
  { k: 'booking', field: 'booking', en: /\b(activat\w*|when can .{0,30}(book|start|use|travel|redeem)|start (using|booking)|notice|in advance|earliest|check.?in|travel (date|period)|booking (period|window|date|starts?)|days before|how soon|redeem\w*)\b/, ar: /(تفعيل|يتفعل|هيتفعل|بيتفعل|يبدا يحجز|امتي يقدر يحجز|يحجز امتي|يستخدمه امتي|اشعار|قبلها ب|قبل السفر|اقرب ميعاد|تشيك ان|ميعاد الحجز)/, words: /activat|notice|in advance|earliest|check.?in|from \d|starting|booking date/i },
  { k: 'others', field: 'others', en: /\b(for (others|someone|somebody|another|other people|his|her|their|family|friends?|wife|husband|spouse|son|daughter|kids|children|parents|mother|mom|father|dad|brother|sister|relatives?|colleague)|someone else|other (people|person|guests?|names?|travell?ers?)|on behalf|transfer\w* (it|the|to)|share (it|the) with)\b/, ar: /(لحد تاني|لحد غيره|لغيره|لغيرها|لاهله|لعيلته|لمراته|لجوزها|لصاحبه|لاصحابه|لاخوه|لاخته|لابنه|لبنته|لوالده|لوالدته|باسم حد|لشخص تاني|لاشخاص تانيين|يحجز لحد)/, words: /others|family|friends|spouse|children|parents|someone|names?|per member|per trip|maximum of/i },
  { k: 'refund', field: 'refund', en: /\b(refund(ed|ing)? (the|their|his|her|it|amount|deposit|money)|is (it|the (amount|deposit|money|card|pass|voucher)) refundable|money back|get .{0,15}money back|withdraw|cash ?out|non.?transferr?able|cancel (the|their|his|her) (purchase|deposit|card|pass|voucher|membership))\b/, ar: /(استرجاع|استرداد|يرجع فلوس|فلوسه ترجع|ترجع الفلوس|يسترد|يسحب الفلوس|فلوسه هترجع|الفلوس (بترجع|هترجع|ترجع|بتترد|تترد)|ريفاند(?! ?بل)|ينفع يرجعها|يرجعها)/, words: /\brefund(?!able (hotel|room|booking|stay|rate))|non.?transfer|monetary value|\bcash\b|withdraw/i },
  { k: 'quote', field: 'quote', en: /\b(quot\w*|show (the )?discount|discount (shows?|appear|visible))\b/, ar: /(عرض سعر|الكوتيشن|كوتيشن|الكوته)/, words: /quot/i },
  { k: 'restr', field: 'restr', en: /\b(restrict\w*|rules?|limitation\w*|exclu\w*|not allowed|can.?t (book|use)|cannot (book|use)|airlines?|pay at (the )?property|eligib\w*|blackout|conditions?|catch)\b/, ar: /(قيود|ممنوع|مستثني|مينفعش|ما ينفعش|الشروط|شروط|طيران معين)/, words: /not|unable|only|exclud|must|cannot|can't|restrict/i },
  { k: 'tc', field: 'tc', en: /\b(t ?& ?c'?s?|terms|t and c|link|url|website|page)\b/, ar: /(الشروط والاحكام|لينك|رابط)/ },
  { k: 'buy', field: 'buy', en: /\b(can .{0,20}(still )?(buy|purchase|join|subscribe|sign up)|available to (buy|purchase)|buying availability|still available|where (to|can .{0,15}) (buy|purchase)|how (to|do .{0,15}) (buy|purchase|join))\b/, ar: /(يشتري|يشترك|ينفع يشتري|متاح للشراء|لسه متاح|يقدر يشتري)/, words: /available|purchase|buy|until/i },
  { k: 'payment', en: /\b(egp|gbp|usd|aed|pounds?|sterling|currenc\w*|pay(ment|ing)?|bank transfer|stripe|credit card|debit card|fees?|instal+ments?|visa|mastercard|exchange rate)\b/, ar: /(جنيه|مصري|استرليني|دولار|عمله|دفع|يدفع|تحويل بنكي|فيزا|رسوم|تقسيط|كاش)/, words: /egp|gbp|usd|pay|bank|stripe|card|fee|currency|exchange|transfer/i },
  { k: 'confidential', en: /\b(confidential\w*|secret|share (it|this|the offer)|can (i|we) (share|send|post)|public|vip)\b/, ar: /(سري|ينفع ابعته|ينفع اقول|خاص|vip)/, words: /confidential|strictly|not to be shared|selected members|vip|keep this/i },
  { k: 'howto', en: /\b(how (does|do) .{0,30}work|how to use|how it works|steps?|process|procedure)\b/, ar: /(بيشتغل ازاي|ازاي بيشتغل|ازاي يستخدم|الخطوات|اشرح)/, words: /how|step|deposit|book|contact|send/i },
  { k: 'benefit', field: 'benefit', en: /\b(how many|discount|how much off|percent\w*|%|benefit|what (do|does|will) .{0,20}get|value|bonus|extra|credit|double|kudos|cashback|reward)\b/, ar: /(خصم|نسبه|كام في الميه|بياخد ايه|هياخد ايه|الميزه|بيستفيد|زياده|ضعف|مكافاه|كودوس)/, words: /\d+ ?%|double|free (room|flight|night|stay|ticket)|extra|bonus/i },
  { k: 'summary', en: /\b(what is|what's|whats|tell me (about|more)|details?|info(rmation)?|summary|overview|key points|explain|about)\b/, ar: /(ايه هو|يعني ايه|تفاصيل|معلومات|عرفني|قولي عن|كلمني عن)/ }
];
const COVER_Q = [
  { i: [0, 1], en: /\b(hotels?|rooms?|stays?|accommodation|resorts?|apartments?|villas?|airbnb)\b/, ar: /(فندق|فنادق|اوتيل|غرفه|غرف|اقامه|شقه|شاليه)/ },
  { i: [1], en: /\b(non.?refundable|nr)\b/, ar: /(غير (قابل|قابله) للاسترداد|نون ريفاندبل|مش ريفاندبل)/ },
  { i: [0], en: /\b(?<!non.?)refundable (hotels?|rooms?|bookings?|stays?)\b/, ar: /(ريفاندبل|قابل للاسترداد)/ },
  { i: [2, 3], en: /\b(flights?|tickets?|air ?(line|fare)s?|fly|plane)\b/, ar: /(طيران|طياره|تذكره|تذاكر|رحلات طيران)/ },
  { i: [3], en: /\b(domestic|local|internal|inside (egypt|the country)) ?(flights?)?\b/, ar: /(داخلي|داخل مصر)/ },
  { i: [2], en: /\b(international|abroad|overseas) ?(flights?)?\b/, ar: /(دولي|خارجي|برا مصر|بره مصر)/ },
  { i: [4], en: /\b(attractions?|activit\w*|transfers?|trains?|ferr(y|ies)|tours?|experiences?|excursions?|cruises?|coach\w*)\b/, ar: /(انشطه|نشاط|فسح|مزارات|مواصلات|قطار|قطر|عبارات|عباره|جولات|رحلات بحريه|كروز)/ },
  { i: [5], en: /\b(theme ?parks?|car ?(rental|hire)s?|rent(ing)? a car|musicals?|disney\w*)\b/, ar: /(ملاهي|عربيه|سياره|ايجار عربيه|تاجير سيارات|ديزني)/ }
];
const COVER_ASK = { en: /\b(cover\w*|use (it|this|the \w+|them)? ?(for|on)|valid (for|on)|apply (to|for|on)|appl(y|ies|icable)|include\w*|work (for|on|with)|eligible|can .{0,25}(book|use))\b/, ar: /(ينفع|يغطي|بيغطي|يشمل|بيشمل|يستخدمه في|يستخدم في|صالح ل|يسري)/ };
const LIST_Q = { en: /\b(which|what|any|list|all|show)\b.{0,40}\b(offers?|products?|promotions?|promos?|deals?|vouchers?|passes|cards?|ones)\b|\b(offers?|products?|promotions?|deals?) (that|which|for|with|covering|on|in|paid in|payable in|accepting)\b/, ar: /(انهي عرض|انهي عروض|ايه العروض|اي عروض|العروض اللي|عروض|انهي منتج|المنتجات)/ };
const NOW_Q = { en: /\b(now|current\w*|active|running|on sale|available (now|today)|live|today|at the moment|ongoing)\b/, ar: /(دلوقتي|حاليا|الحاليه|الحالي|شغال|شغاله|النهارده|متاح)/ };
const NEW_Q = { en: /\b(latest|newest|new|recent\w*|last (offer|promo|product)\w*|just (launched|released))\b/, ar: /(جديد|احدث|اخر عرض|اخر العروض|الجديده|نزل)/ };

const L = {
  min: ['Minimum purchase', 'الحد الأدنى للشراء'], max: ['Maximum amount', 'الحد الأقصى'], expiry: ['Expiry', 'الصلاحية'], offer: ['Offer period (when it can be bought)', 'مدة العرض (فترة الشراء)'],
  booking: ['Booking / activation period', 'فترة الحجز / التفعيل'], others: ['Booking for others', 'الحجز لأشخاص آخرين'], refund: ['Is the amount refundable?', 'هل المبلغ قابل للاسترداد؟'],
  quote: ['Does the quotation show the discount?', 'هل عرض السعر يظهر الخصم؟'], restr: ['Booking restrictions', 'قيود الحجز'], tc: ['Terms & Conditions', 'الشروط والأحكام'],
  buy: ['Buying availability', 'إتاحة الشراء'], benefit: ['Benefit / discount', 'الميزة / الخصم'], payment: ['Payment & currency', 'الدفع والعملة'],
  confidential: ['Sharing / confidentiality', 'السرية والمشاركة'], howto: ['How it works', 'طريقة الاستخدام'], cover: ['What it can be used for', 'تستخدم في إيه'],
  also: ['Also in this offer', 'مذكور كمان في العرض'], versions: ['Other versions', 'نسخ تانية من نفس العرض'], status: ['Status', 'الحالة'], notFound: ['Not in the knowledge base', 'مش موجود في قاعدة المعرفة']
};

/* ---------- index ---------- */
export function buildIndex(kbItems, updates = []) {
  // the same offer can sit in both the Promotions and Products sheets: keep the fuller copy
  const byName = new Map();
  for (const it of kbItems) { const key = tokens(it.name).join(' '); const prev = byName.get(key); if (!prev || Object.keys(it.fields).length > Object.keys(prev.fields).length) byName.set(key, it); }
  const items = [...byName.values()].map(it => {
    const f = {}; for (const [k, v] of Object.entries(it.fields)) { const role = FIELDS.find(([, rx]) => rx.test(k)); if (role && !f[role[0]]) f[role[0]] = { label: k, text: v }; }
    const nameT = tokens(it.name); const core = [...new Set(nameT.filter(t => !STOP.has(t) && !NAME_NOISE.has(t) && !/^\d+$/.test(t) && !(t in MONTHS) && t.length > 1))];
    const all = Object.entries(it.fields).filter(([k]) => !/promotion details/i.test(k));
    const sents = []; for (const [k, v] of all) for (const s of sentences(v)) sents.push({ field: k, s, t: tokens(s) });
    return { it, id: it.id, name: it.name, f, nameT, core, sents, text: all.map(([, v]) => v).join('\n') };
  });
  // recency: the offer's end date, or the nearest neighbour's for undated entries (sheets are in date order)
  const byCat = {}; items.forEach(x => (byCat[x.it.cat] = byCat[x.it.cat] || []).push(x));
  Object.values(byCat).forEach(list => { list.sort((a, b) => b.it.rank - a.it.rank); list.forEach((x, i) => { if (x.it.end) x.when = Date.parse(x.it.end); else { const n = [...list.slice(0, i).reverse(), ...list.slice(i + 1)].find(y => y.it.end); x.when = n ? Date.parse(n.it.end) + (i - list.indexOf(n)) : 0; } }); });
  const df = {}; items.forEach(x => x.core.forEach(t => df[t] = (df[t] || 0) + 1)); const N = items.length;
  const idf = t => Math.log(1 + N / (df[t] || 0.5));
  const vocab = Object.keys(df);
  // sentence-level index for deep search (BM25)
  const docs = []; items.forEach(x => x.sents.forEach(s => docs.push({ x, ...s })));
  (updates || []).forEach(u => sentences(u.Body).concat([u.Title]).forEach(s => docs.push({ x: null, upd: u, field: 'Team update', s, t: tokens(s) })));
  const sdf = {}; docs.forEach(d => new Set(d.t).forEach(t => sdf[t] = (sdf[t] || 0) + 1)); const avg = docs.reduce((a, d) => a + d.t.length, 0) / (docs.length || 1);
  return { items, idf, df, vocab, docs, sdf, avg, N, updates };
}

/* ---------- find the offer(s) meant ---------- */
function resolve(ix, q, qT) {
  const qn = ' ' + tokens(q).join(' ') + ' ';
  // exact full-name mention wins (used by follow-up buttons and [[links]])
  const exact = ix.items.filter(x => x.nameT.length && qn.includes(' ' + x.nameT.join(' ') + ' ')).sort((a, b) => b.nameT.length - a.nameT.length);
  // (a bare product name like "Magic Pass" means the newest version, so only names carrying a date/version count as exact)
  const ex = exact.find(x => x.nameT.length >= 2 && (x.nameT.length > x.core.length + 1 || !siblings(ix, x).some(y => y.when > x.when)));
  if (ex) return { best: ex, score: 99, exact: true, named: true, hits: ex.nameT, versions: siblings(ix, ex) };
  const qCore = qT.filter(t => !STOP.has(t)); const nums = qT.filter(t => /^\d+$/.test(t)); const mon = qT.map(t => MONTHS[t]).filter(Boolean);
  const scored = [];
  for (const x of ix.items) {
    if (!x.core.length) continue;
    let w = 0, tot = 0, rare = false; const used = new Set();
    let named = false; const hits = [];
    for (const c of x.core) { const g = GENERIC.has(c) ? 0.35 : 1; tot += ix.idf(c) * g; const hit = qCore.find(t => !used.has(t) && close(t, c)); if (hit) { used.add(hit); hits.push(hit); w += ix.idf(c) * g; if (g === 1) named = true; if ((ix.df[c] || 0) <= 4 && g === 1) rare = true; } }
    if (!w) continue;
    const cov = w / tot; let s = w * (0.3 + 0.7 * cov * cov) + (cov > 0.99 ? 0.8 : 0);
    // version hints: month / year / day numbers in the question that also appear in the name
    const nmNums = x.nameT.filter(t => /^\d+$/.test(t)); const nmMon = x.nameT.map(t => MONTHS[t]).filter(Boolean);
    let numHit = false; nums.forEach(n => { if (nmNums.some(m => m === n || m === n.slice(-2) || n === m.slice(-2))) { s += n.length >= 2 && +n >= 10 && +n <= 100 ? 2.4 : 1.2; numHit = true; } });
    if (numHit && cov > 0.99) { named = true; s += 2; }
    mon.forEach(m => { if (nmMon.includes(m) || nmNums.includes(String(m).padStart(2, '0'))) s += 1.2; });
    scored.push({ x, s, cov, w, rare, named, hits, numHit });
  }
  if (!scored.length) return null;
  scored.sort((a, b) => b.s - a.s || b.x.when - a.x.when);
  const top = scored[0];
  // every product word the agent typed was found in this offer's name
  const qVocab = qCore.filter(t => !GENERIC.has(t) && ix.vocab.some(v => close(t, v)));
  const allUsed = qVocab.length > 0 && qVocab.every(t => top.hits.includes(t));
  const strong = (top.cov > 0.99 && (top.x.core.length >= 2 || top.numHit) && top.w >= 0.3) || (top.named && allUsed && top.cov >= (top.numHit ? 0.3 : 0.45)) || top.named && ((top.cov >= 0.74 && top.w >= 1.4) || (top.rare && top.cov >= 0.34) || top.s >= 6);
  if (!strong) return { weak: scored.slice(0, 5).map(s => s.x) };
  // among equally good matches (same product, different dates) take the newest unless the question named a version
  const near = scored.filter(s => s.s >= top.s - 0.01 && s.cov >= top.cov - 0.01);
  const best = near.sort((a, b) => b.x.when - a.x.when)[0].x;
  const vers = siblings(ix, best);
  return { best, score: top.s, named: top.named, rare: top.rare, hits: scored.find(s => s.x === best).hits, versions: vers, close: scored.filter(s => s.x !== best && !vers.includes(s.x) && s.s >= top.s * 0.9 && s.x.core.join(' ') !== best.core.join(' ')).slice(0, 3).map(s => s.x) };
}
function siblings(ix, best) {
  const key = best.core.join(' ');
  return ix.items.filter(x => x !== best && x.core.length && (x.core.join(' ') === key || (x.core.length >= 2 && best.core.length >= 2 && x.core.every(c => best.core.includes(c)) && best.core.every(c => x.core.includes(c))))).sort((a, b) => b.when - a.when).slice(0, 4);
}

/* ---------- deep search across every sentence ---------- */
function search(ix, qT, within, opts = {}) {
  const terms = [...new Set(qT.filter(t => !STOP.has(t) && t.length > 1))]; if (!terms.length) return [];
  const docs = within ? ix.docs.filter(d => d.x === within) : ix.docs; const Nd = ix.docs.length; const out = [];
  for (const d of docs) {
    let s = 0; const tf = {}; d.t.forEach(t => tf[t] = (tf[t] || 0) + 1);
    for (const q of terms) { const hits = Object.keys(tf).filter(t => close(q, t)); if (!hits.length) continue; const f = hits.reduce((a, t) => a + tf[t], 0); const n = Math.max(...hits.map(t => ix.sdf[t] || 1));
      const idf = Math.log(1 + (Nd - n + 0.5) / (n + 0.5)); s += idf * (f * 2.2) / (f + 1.2 * (0.25 + 0.75 * d.t.length / ix.avg)); }
    if (opts.words && opts.words.test(d.s)) s += 1.5;
    if (s > 0) out.push({ d, s });
  }
  return out.sort((a, b) => b.s - a.s);
}
function bold(s, qT) { const terms = qT.filter(t => !STOP.has(t) && t.length > 2); if (!terms.length) return s; return s.replace(/[A-Za-z؀-ۿ0-9£%]+/g, w => terms.some(t => close(t, stem(normAr(w.toLowerCase())))) ? `**${w}**` : w); }

/* ---------- writing the answer ---------- */
function statusLine(x, ar) {
  const it = x.it; const d = it.end ? new Date(it.end).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : '';
  const m = {
    'On sale': [`🟢 On sale until ${d}.`, `🟢 متاح للشراء حتى ${d}.`],
    'Ended': [`⚪ Sale ended ${d} — members who already bought it keep their benefits under its T&Cs.`, `⚪ العرض انتهى ${d} — اللي اشترى قبل كده بيستفيد حسب الشروط والأحكام.`],
    'While quota lasts': ['🟡 Sold while the allocated quota lasts.', '🟡 متاح لحد ما الكمية المخصصة تخلص.'],
    'Always available': ['🔵 Available all year.', '🔵 متاح طول السنة.'], 'Paused': ['🔴 Paused — check the offer period.', '🔴 متوقف حاليًا — راجع مدة العرض.']
  }[it.status];
  return m ? m[ar ? 1 : 0] : '';
}
function block(text) { const lines = sentences(text); if (lines.length <= 1) return clean(text).replace(/\s*\n\s*/g, ' '); return lines.map(l => '- ' + l).join('\n'); }
function confidentialNote(x, ar) { return /confidential|strictly not|not to be shared|keep this confidential|selected members only|vip list/i.test(x.text) ? (ar ? '\n⚠️ **العرض ده سري / لأعضاء محددين — ما يتشاركش بشكل عام.**' : '\n⚠️ **This offer is confidential / for selected members only — do not share it publicly.**') : ''; }
function tail(x, ar, res, opts = {}) {
  const out = [''];
  const st = statusLine(x, ar); if (st) out.push(st);
  const c = confidentialNote(x, ar); if (c) out.push(c);
  if (opts.tc !== false && x.it.link) out.push((ar ? 'الشروط والأحكام: ' : 'T&Cs: ') + x.it.link);
  const v = (res && res.versions || []).filter(y => y !== x); if (v.length && !res.exact) out.push('', (ar ? L.versions[1] : L.versions[0]) + ': ' + v.map(y => `[[${y.name}]]`).join(' · '));
  const c2 = (res && res.close || []).filter(y => y !== x); if (c2.length && !res.exact) out.push((v.length ? '' : '\n') + (ar ? 'ممكن تقصد كمان: ' : 'You may also mean: ') + c2.map(y => `[[${y.name}]]`).join(' · '));
  return out.join('\n');
}
function coverText(x, which, ar) {
  return which.map(i => { const c = COVER[i]; const v = x.f[c.k] ? x.f[c.k].text : ''; const y = yn(v);
    const mark = y === true ? '✅' : y === false ? '❌' : y === 'partial' ? '🟡' : '❔'; const word = y === true ? (ar ? 'نعم' : 'Yes') : y === false ? (ar ? 'لا' : 'No') : y === 'partial' ? (ar ? 'جزئيًا' : 'Partly') : (ar ? 'غير مذكور' : 'Not stated');
    const extra = v && !/^(yes|no|x|[-_—.]+)$/i.test(clean(v)) ? ` — ${cut(v, 160)}` : '';
    return `- ${mark} ${ar ? c.ar : c.en}: **${word}**${extra}`; }).join('\n');
}
function related(x, intent, primaryText, qT, max = 3) {
  if (!intent.words) return [];
  const seen = new Set(sentences(primaryText).map(s => s.toLowerCase()));
  return x.sents.filter(s => intent.words.test(s.s) && !seen.has(s.s.toLowerCase()) && (!primaryText || !primaryText.includes(s.s)) && !/^(yes|no|x)$/i.test(s.s))
    .map(s => ({ s, w: (intent.field && x.f[intent.field] && s.field === x.f[intent.field].label ? -1 : 0) + qT.filter(t => s.t.includes(t)).length })).sort((a, b) => b.w - a.w).slice(0, max).map(r => `- ${r.s.s} _(${r.s.field})_`);
}
function summary(x, ar, res) {
  const g = k => x.f[k] && clean(x.f[k].text);
  const lines = [`**[[${x.name}]]** · ${x.it.cat}`];
  if (g('benefit')) lines.push('', block(cut(x.f.benefit.text, 900)));
  const facts = [['offer', L.offer], ['booking', L.booking], ['expiry', L.expiry], ['min', L.min], ['others', L.others], ['refund', L.refund], ['quote', L.quote]]
    .filter(([k]) => g(k) && !/^[-_—]+$/.test(g(k))).map(([k, l]) => `- **${ar ? l[1] : l[0].replace(/ \(.*\)/, '')}:** ${/^x$/i.test(g(k)) ? (ar ? 'لا' : 'No') : cut(x.f[k].text, 190)}`);
  if (facts.length) lines.push('', ...facts);
  const yes = COVER.filter(c => [true, 'partial'].includes(yn(x.f[c.k] && x.f[c.k].text))).map(c => (ar ? c.ar : c.en) + (yn(x.f[c.k].text) === 'partial' ? (ar ? ' (جزئيًا)' : ' (partly)') : '')), no = COVER.filter(c => yn(x.f[c.k] && x.f[c.k].text) === false).map(c => ar ? c.ar : c.en);
  if (yes.length || no.length) lines.push('', `**${ar ? L.cover[1] : L.cover[0]}:** ${yes.length ? '✅ ' + yes.join(', ') : ''}${yes.length && no.length ? '  ·  ' : ''}${no.length ? '❌ ' + no.join(', ') : ''}`);
  if (g('restr')) lines.push('', `**${ar ? L.restr[1] : L.restr[0]}:**`, block(cut(x.f.restr.text, 600)));
  return lines.join('\n') + tail(x, ar, res);
}
function fieldAnswer(x, intent, ar, qT, res) {
  const fld = intent.field && x.f[intent.field]; const lab = ar ? L[intent.k][1] : L[intent.k][0];
  const val = fld ? clean(fld.text) : '';
  const out = [`**${lab}** — [[${x.name}]]`];
  if (val && !/^[-_—]+$/.test(val)) {
    const y = ['others', 'refund', 'quote'].includes(intent.k) ? yn(val) : null;
    if (y === true) out.push(ar ? '**نعم.**' : '**Yes.**'); else if (y === false) out.push(ar ? '**لا.**' : '**No.**'); else if (y === 'unsure') out.push(ar ? '**مش مؤكد في قاعدة المعرفة** — اتأكد من التيم ليدر.' : '**Not confirmed in the knowledge base** — check with your team leader.');
    if (!/^(yes|no|x|not sure)\.?$/i.test(val)) out.push(block(val));
  } else out.push(ar ? '_الخانة دي مش متعبية للعرض ده في الشيت._' : '_The sheet doesn\'t fill this in for this offer._');
  const rel = related(x, intent, val, qT); if (rel.length) out.push('', `**${ar ? L.also[1] : L.also[0]}:**`, ...rel);
  return { text: out.join('\n') + tail(x, ar, res, { tc: intent.k !== 'tc' }), conf: val && !/^[-_—]+$/.test(val) ? 0.92 : (rel.length ? 0.7 : 0.45) };
}
function wordsAnswer(x, intent, ar, qT, res) {
  const lab = ar ? L[intent.k][1] : L[intent.k][0];
  const hits = x.sents.filter(s => intent.words.test(s.s)).map(s => ({ s, w: qT.filter(t => s.t.includes(t)).length + (/restrict|benefit|discount|min/i.test(s.field) ? 0.5 : 0) })).sort((a, b) => b.w - a.w).slice(0, 5);
  const out = [`**${lab}** — [[${x.name}]]`];
  if (hits.length) out.push(...hits.map(h => `- ${bold(h.s.s, qT)} _(${h.s.field})_`));
  else if (intent.k === 'confidential') out.push(ar ? 'الشيت **مش مكتوب فيه إن العرض ده سري** أو لأعضاء محددين. لو مش متأكد ارجع للتيم ليدر قبل ما تشاركه.' : 'The sheet **does not mark this offer as confidential** or for selected members only. If in doubt, check with your team leader before sharing it.');
  else out.push(ar ? '_مفيش ذكر لده في بيانات العرض._' : '_Nothing about this is mentioned for this offer._');
  return { text: out.join('\n') + tail(x, ar, res), conf: hits.length || intent.k === 'confidential' ? 0.8 : 0.4 };
}
function listLine(x, ar) {
  const st = { 'On sale': ar ? 'متاح' : 'on sale', 'Ended': ar ? 'انتهى' : 'sale ended', 'While quota lasts': ar ? 'لحد ما الكمية تخلص' : 'while quota lasts', 'Always available': ar ? 'طول السنة' : 'all year', 'Paused': ar ? 'متوقف' : 'paused' }[x.it.status] || '';
  const d = x.it.end ? new Date(x.it.end).toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' }) : '';
  return `- [[${x.name}]] — ${st}${d && st ? ' ' + d : ''}${x.f.benefit ? ' · ' + cut(x.f.benefit.text, 90) : ''}`;
}
const FOLLOW = [['benefit', 'What is the benefit?', 'إيه الميزة؟'], ['min', 'Minimum purchase', 'أقل مبلغ'], ['expiry', 'Expiry', 'الصلاحية'], ['booking', 'Activation & notice', 'التفعيل'], ['others', 'Booking for others', 'الحجز لحد تاني'],
  ['cover', 'What does it cover?', 'بيغطي إيه؟'], ['restr', 'Booking restrictions', 'قيود الحجز'], ['refund', 'Is it refundable?', 'الفلوس بترجع؟'], ['summary', 'Full summary', 'ملخص كامل']];
function follows(x, done, ar) { return FOLLOW.filter(([k]) => k !== done && (k === 'cover' || k === 'summary' || x.f[k])).slice(0, 5).map(([k, en, a]) => ({ label: ar ? a : en, q: `${cut(x.name, 50)} — ${ar ? a : en}`, item: x.id })); }

/* ---------- main ---------- */
export function answer(ix, question, opts = {}) {
  const raw = String(question || '').trim(); const ar = isArabic(raw);
  const lo = ' ' + raw.toLowerCase().replace(/[’']/g, "'") + ' '; const la = ' ' + normAr(raw.toLowerCase()) + ' ';
  const m = r => (r.en && r.en.test(lo)) || (r.ar && r.ar.test(la));
  const qT = tokens(raw);
  const ctx = opts.ctx ? ix.items.find(x => x.id === opts.ctx) : null;
  const forced = opts.item ? ix.items.find(x => x.id === opts.item) : null;
  const res = forced ? { best: forced, exact: true, versions: siblings(ix, forced) } : (resolve(ix, raw, qT) || {});
  // the question names no offer: follow on from the previous answer if it reads like a follow-up
  let x = res.best || null; let viaCtx = false;
  // a follow-up ("is it refundable?", "can he use it for flights?") stays on the previous offer unless a new product is clearly named
  const pron = /\b(it|its|this|that|these|those|the same|this one|that one|the offer|the pass|the card|the voucher|he|she|they)\b/.test(lo) || /(ده|دي|دا|ديه|نفسه|نفس العرض|بتاعه|بتاعها)/.test(la);
  if (ctx && !forced && x && ((!res.exact && (pron || !res.rare) && !(res.hits || []).some(h => !GENERIC.has(h) && !ctx.core.some(c => close(h, c)) && x.core.some(c => close(h, c)))))) {
    if (x !== ctx && (pron || !res.named || (res.hits || []).every(h => GENERIC.has(h)))) { x = ctx; viaCtx = true; }
  }
  // look for the question type in what is left once the offer's own name is taken out ("Travel Pounds" is not a currency question)
  let lo2 = lo, la2 = la;
  if (forced && raw.includes(' — ')) { const tailQ = raw.slice(raw.lastIndexOf(' — ') + 3); lo2 = ' ' + tailQ.toLowerCase() + ' '; la2 = ' ' + normAr(tailQ.toLowerCase()) + ' '; }
  if (x) { const esc = t => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); const full = new RegExp(esc(x.name.toLowerCase().trim()).replace(/\s+/g, '\\s+'), 'g'); lo2 = lo2.replace(full, ' '); la2 = la2.replace(full, ' '); }
  if (x) { const nameWords = x.core.concat(x.nameT); const strip = s => s.replace(/[A-Za-z\u0600-\u06ff0-9']+/g, w => { const t = tokens(w); return t.length && t.every(u => nameWords.some(c => c.length > 2 && close(u, c))) ? ' ' : w; }); lo2 = strip(lo2); la2 = strip(la2); }
  const m2 = r => (r.en && r.en.test(lo2)) || (r.ar && r.ar.test(la2));
  const intents = INTENTS.filter(i => m2(i)).map(i => i.k);
  if (/\brefundable\b(?! (hotels?|rooms?|bookings?|stays?|rates?))/.test(lo2) && !/non.?refundable/.test(lo2) && !intents.includes('refund')) intents.unshift('refund');
  let coverWhich = [...new Set(COVER_Q.filter(c => m2(c)).flatMap(c => c.i))];
  // a specific kind wins over the general word ("domestic flights" ≠ all flights)
  const spec = k => m2(COVER_Q[k]); if (spec(1)) coverWhich = coverWhich.filter(i => i !== 0); else if (spec(2)) coverWhich = coverWhich.filter(i => i !== 1);
  if (spec(4) && !spec(5)) coverWhich = coverWhich.filter(i => i !== 2); else if (spec(5) && !spec(4)) coverWhich = coverWhich.filter(i => i !== 3);
  if (intents.includes('refund') && !m2(COVER_ASK)) coverWhich = [];
  const offerWord = /\b(offers?|deals?|promos?|promotions?|products?)\b/.test(lo) || /(عروض|العروض)/.test(la);
  const plural = /\b(offers|products|promotions|deals|promos|ones|vouchers|passes|cards)\b/.test(lo) || /(عروض|العروض|منتجات)/.test(la);
  const isList = (m(LIST_Q) && (!x || plural)) || ((m(NOW_Q) || m(NEW_Q)) && !x) || (!x && offerWord);
  // list questions read the whole sentence (an offer's name may contain "non-refundable hotels")
  const coverFor = mm => { let w = [...new Set(COVER_Q.filter(c => mm(c)).flatMap(c => c.i))]; const sp = k => mm(COVER_Q[k]);
    if (sp(1)) w = w.filter(i => i !== 0); else if (sp(2)) w = w.filter(i => i !== 1); if (sp(4) && !sp(5)) w = w.filter(i => i !== 2); else if (sp(5) && !sp(4)) w = w.filter(i => i !== 3); return w; };
  if (isList) coverWhich = coverFor(m);
  // "what about the e& version?" → the previous offer's name + the new detail
  if (ctx && (!x || x === ctx) && !forced) { const extra = qT.filter(t => !STOP.has(t) && !GENERIC.has(t) && !ctx.core.some(c => close(t, c)) && ix.vocab.some(v => close(t, v)));
    if (extra.length) { const r2 = resolve(ix, ctx.core.join(' ') + ' ' + raw, tokens(ctx.core.join(' ') + ' ' + raw)); if (r2 && r2.best && r2.best !== ctx && extra.every(t => r2.best.core.some(c => close(t, c)))) { x = r2.best; viaCtx = false; Object.assign(res, r2); } } }
  if (!x && ctx && !isList && (intents.length || coverWhich.length || pron || (search(ix, qT, ctx)[0] || {}).s > 2.5)) { x = ctx; viaCtx = true; }
  const out = (text, conf, refs = [], follow = []) => ({ answer: text, conf, refs, follow, lang: ar ? 'ar' : 'en', item: x ? x.id : null });

  // 1) comparison of two offers
  if (intents.includes('compare')) {
    const parts = raw.split(/\bvs\.?\b|\bversus\b|\band\b|\bor\b|\bwith\b|,|\bو|\bولا\b|\bمع\b/i).map(s => s.replace(/\b(compare|difference|between|what is the|what's the)\b|الفرق بين|قارن بين|قارن/gi, '').trim()).filter(s => tokens(s).length);
    const found = []; parts.forEach(p => { const r = resolve(ix, p, tokens(p)); if (r && r.best && !found.includes(r.best)) found.push(r.best); });
    if (found.length >= 2) {
      const [a, b] = found; const rows = [['benefit', L.benefit], ['offer', L.offer], ['booking', L.booking], ['expiry', L.expiry], ['min', L.min], ['others', L.others], ['refund', L.refund]];
      const t = [`**${ar ? 'مقارنة' : 'Comparison'}: [[${a.name}]] vs [[${b.name}]]**`];
      rows.forEach(([k, l]) => { if (!a.f[k] && !b.f[k]) return; t.push('', `**${ar ? l[1] : l[0].replace(/ \(.*\)/, '')}**`, `- ${cut(a.name, 40)}: ${a.f[k] ? cut(a.f[k].text, 220) : '—'}`, `- ${cut(b.name, 40)}: ${b.f[k] ? cut(b.f[k].text, 220) : '—'}`); });
      t.push('', `**${ar ? L.cover[1] : L.cover[0]}**`); COVER.forEach(c => { const ya = yn(a.f[c.k] && a.f[c.k].text), yb = yn(b.f[c.k] && b.f[c.k].text); t.push(`- ${ar ? c.ar : c.en}: ${ya === true ? '✅' : ya === false ? '❌' : '❔'} / ${yb === true ? '✅' : yb === false ? '❌' : '❔'}`); });
      return out(t.join('\n'), 0.85, [a.id, b.id]);
    }
  }

  // 2) list questions: "which offers cover NR hotels", "what's on sale now", "latest offers"
  if (isList || (!x && (coverWhich.length || m(NOW_Q) || m(NEW_Q)))) {
    let list = ix.items.slice(); const why = [];
    if (coverWhich.length) { list = list.filter(y => coverWhich.some(i => [true, 'partial'].includes(yn(y.f[COVER[i].k] && y.f[COVER[i].k].text)))); why.push((ar ? 'تغطي ' : 'covering ') + coverWhich.map(i => ar ? COVER[i].ar : COVER[i].en).join(ar ? ' أو ' : ' or ')); }
    const wantNow = m(NOW_Q);
    if (wantNow) { list = list.filter(y => ['On sale', 'While quota lasts', 'Always available'].includes(y.it.status)); why.push(ar ? 'متاحة دلوقتي' : 'available now'); }
    // remaining topic words (e.g. "EGP", "HSBC", "Umrah") narrow the list through the offer's text
    const topic = qT.filter(t => !STOP.has(t) && !/^(offer|product|promotion|promo|deal|which|list|show|now|current|active|available|latest|newest|new|recent|cover|use|valid|include|hotel|flight|ticket|room|refundable|non|nr|international|domestic|attraction|theme|park|car|rental|sale|today|running|live|use|used|using|book|booking|covered|covering|valid|apply|eligible|include|includ|there|anything|something|offers|promotions)$/.test(t) && t.length > 2);
    if (topic.length) { const tl = list.filter(y => topic.every(t => y.core.some(c => close(t, c)) || y.sents.some(s => s.t.some(w => close(t, w))))); if (tl.length) { list = tl; why.push((ar ? 'فيها ' : 'mentioning ') + topic.join(', ')); } }
    list.sort((a, b) => b.when - a.when);
    if (!list.length) {
      const t = wantNow ? (ar ? 'مفيش عروض متاحة للشراء دلوقتي حسب التواريخ اللي في الشيت. أحدث العروض:' : 'By the dates in the sheet, no offer is on sale right now. The most recent ones are:') : (ar ? 'مالقيتش عروض بالشرط ده.' : 'I couldn\'t find offers matching that.');
      const recent = ix.items.slice().sort((a, b) => b.when - a.when).slice(0, 6);
      return out(t + '\n' + recent.map(y => listLine(y, ar)).join('\n'), 0.6, recent.map(y => y.id));
    }
    const plainNew = m(NEW_Q) && !why.length; if (plainNew) list = list.slice(0, 8);
    const shown = list.slice(0, 10);
    const head = plainNew ? (ar ? '**أحدث العروض:**' : '**Latest offers (newest first):**') : ar ? `**${list.length} عرض**${why.length ? ' ' + why.join('، ') : ''} (الأحدث الأول):` : `**${list.length} ${list.length === 1 ? 'offer' : 'offers'}**${why.length ? ' ' + why.join(', ') : ''} (newest first):`;
    return out([head, ...shown.map(y => listLine(y, ar)), list.length > 10 ? (ar ? `\nو ${list.length - 10} كمان — استخدم الفلاتر في صفحة المعرفة.` : `\n…and ${list.length - 10} more — use the filters on the Knowledge page to see them all.`) : ''].join('\n'), 0.85, shown.map(y => y.id));
  }

  // 3) a specific offer
  if (x) {
    const r = viaCtx ? { versions: [] } : res; let pick = intents.filter(k => !['summary', 'compare', 'howto'].includes(k));
    if ((intents.includes('summary') || intents.includes('howto')) && pick.every(k => k === 'benefit')) pick = [];
    // coverage questions ("can they use it for flights?")
    if (coverWhich.length && (m2(COVER_ASK) || !pick.length)) {
      const t = [`**${ar ? L.cover[1] : L.cover[0]}** — [[${x.name}]]`, coverText(x, coverWhich, ar)];
      const typeWords = new RegExp(COVER_Q.filter(c => c.i.some(i => coverWhich.includes(i))).map(c => c.en.source).join('|'), 'i');
      const rel = x.sents.filter(s => typeWords.test(s.s) && !COVER.some(c => x.f[c.k] && s.field === x.f[c.k].label)).slice(0, 3).map(s => `- ${s.s} _(${s.field})_`);
      if (rel.length) t.push('', `**${ar ? L.also[1] : L.also[0]}:**`, ...rel);
      return out(t.join('\n') + tail(x, ar, r), 0.9, [x.id], follows(x, 'cover', ar));
    }
    if (pick.length) {
      const k = pick.find(k => k !== 'benefit') || pick[0]; const intent = INTENTS.find(i => i.k === k);
      const a = intent.field ? fieldAnswer(x, intent, ar, qT, r) : wordsAnswer(x, intent, ar, qT, r);
      // if a second, different question type was asked too, add it briefly
      const second = pick.find(j => j !== k && j !== 'benefit' && INTENTS.find(i => i.k === j).field);
      let text = a.text; if (second) { const i2 = INTENTS.find(i => i.k === second); const v2 = x.f[i2.field]; if (v2) text += `\n\n**${ar ? L[second][1] : L[second][0]}:** ${cut(v2.text, 300)}`; }
      return out(text, a.conf, [x.id], follows(x, k, ar));
    }
    // something specific but not one of the standard fields: search inside this offer
    const leftover = qT.filter(t => !STOP.has(t) && !x.core.some(c => close(t, c)) && !x.nameT.includes(t) && !(t in MONTHS) && !/^\d+$/.test(t) && !/^(offer|product|promotion|promo|deal|work|use|explain|about|info|detail|summary|tell|know|need)$/.test(t));
    if (leftover.length && !intents.includes('summary') && !intents.includes('howto')) {
      const hits = search(ix, leftover, x).slice(0, 4).filter(h => h.s > 1);
      if (hits.length) return out([`[[${x.name}]] — ${ar ? 'اللي لقيته في بيانات العرض:' : 'here is what the offer says:'}`, ...hits.map(h => `- ${bold(h.d.s, leftover)} _(${h.d.field})_`)].join('\n') + tail(x, ar, r), 0.75, [x.id], follows(x, '', ar));
    }
    return out(summary(x, ar, r), 0.8, [x.id], follows(x, 'summary', ar));
  }

  // 4) no specific offer: dig through every sentence of every offer and team update
  const hits = search(ix, qT, null, { words: (INTENTS.find(i => intents.includes(i.k) && i.words) || {}).words });
  const byItem = new Map(); for (const h of hits) { const key = h.d.x ? h.d.x.id : 'U' + h.d.upd.ID; if (!byItem.has(key)) byItem.set(key, []); byItem.get(key).push(h); }
  const groups = [...byItem.values()].map(l => ({ l, s: l[0].s + (l[1] ? l[1].s * 0.3 : 0) })).sort((a, b) => b.s - a.s);
  const best = groups[0];
  if (best && best.s >= 3) {
    const lines = [ar ? 'دي أقرب حاجات لقيتها في قاعدة المعرفة:' : 'Here is what I found in the knowledge base:'];
    const refs = [];
    groups.slice(0, 5).filter(g => g.s >= best.s * 0.45).forEach(g => {
      const d = g.l[0].d; const title = d.x ? `[[${d.x.name}]]` : `**${ar ? 'تحديث للفريق' : 'Team update'}: ${d.upd.Title}**`; if (d.x) refs.push(d.x.id);
      lines.push('', title + (d.x && d.x.it.status ? ` · ${listLine(d.x, ar).split(' — ')[1] ? listLine(d.x, ar).split(' — ')[1].split(' · ')[0] : ''}` : ''));
      g.l.slice(0, 2).forEach(h => lines.push(`- ${bold(cut(h.d.s, 320), qT)} _(${h.d.field})_`));
    });
    if (res.weak && res.weak.length) lines.push('', (ar ? 'تقصد عرض معين؟ ' : 'Did you mean a specific offer? ') + res.weak.slice(0, 4).map(y => `[[${y.name}]]`).join(' · '));
    return out(lines.join('\n'), best.s >= 6 ? 0.65 : 0.5, refs, refs.length === 1 ? [] : []);
  }
  const sugg = (res.weak || []).slice(0, 4);
  return out((ar ? `**${L.notFound[1]}.** مالقيتش إجابة للسؤال ده في الشيت — اسأل التيم ليدر أو الديوتي مانجر.` : `**${L.notFound[0]}.** I couldn't find an answer to that in the sheet — please check with your team leader or duty manager.`)
    + (sugg.length ? '\n\n' + (ar ? 'ممكن تقصد: ' : 'Did you mean: ') + sugg.map(y => `[[${y.name}]]`).join(' · ') : '')
    + '\n\n' + (ar ? 'جرب تكتب اسم العرض + سؤالك، مثلاً: "Magic Pass أقل مبلغ"' : 'Tip: include the offer name, e.g. "Magic Pass minimum purchase" or "Flight Club expiry".'), 0.1, sugg.map(y => y.id));
}
export const _t = { tokens, resolve, yn, isArabic, search, INTENTS, COVER_Q };
