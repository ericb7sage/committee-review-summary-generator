const REQUIRED_HEADERS = [
  "Reviewer",
  "File",
  "Why Law?",
  "Thrive?",
  "Contribute?",
  "Know?",
  "Tags",
  "Reach",
  "Target",
  "Safety",
  "Softs",
  "Notes",
  "Anything Else?",
];

const RUBRIC_DELTA = {
  "Strongly Agree": 1,
  Agree: 0,
  Neutral: -1,
  Disagree: -2,
  "Strongly Disagree": -2,
};

const BAND_ORDER = ["T3", "T6", "T14", "T20", "T30", "T50", "T75", "T100", "T100+"];

const BAND_LABELS = {
  T3: "Top 3",
  T6: "Top 6",
  T14: "Top 14",
  T20: "Top 20",
  T30: "Top 30",
  T50: "Top 50",
  T75: "Top 75",
  T100: "Top 100",
  "T100+": "Top 100+",
};

function normalizeTagKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[’‘]/g, "'")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

// Generated from Tag Explanation (2).csv
const TAG_DEFINITIONS = [
  {
    "name": "Likable",
    "polarity": "positive",
    "description": "Admissions committees are building a community, not filling seats. Something about your writing made us feel like we know and like you. Likability is often established through a warm tone, a flash of humor, or a moment of genuine vulnerability. Being likable doesn’t get you into law school—but it might be the slight edge that gets you picked over someone equally qualified."
  },
  {
    "name": "Unique Perspective",
    "polarity": "positive",
    "description": "We can imagine how your background would help you make a unique contribution to a law school class. This can come from cultural diversity or diversity of life experiences. This is particularly helpful for top schools—with so many excellent candidates, T14 admissions officers often consider how you might complement their goal of building a well-rounded class. Law school classrooms work best when the people in them have genuinely different ways of seeing the world."
  },
  {
    "name": "Strong Hire",
    "polarity": "positive",
    "description": "Your profile suggests you'll be a strong candidate on the job market. In recent years, this has become one of the primary things admissions officers search for. Employment outcomes dominate school ranking metrics. You probably have some combination of professional polish, clear career direction, domain expertise, or the kind of résumé that signals to employers that you can hit the ground running."
  },
  {
    "name": "Overcame Challenges",
    "polarity": "positive",
    "description": "We can tell that you’ve overcome disadvantages, be they economic hardship or difficult personal situations. While we never recommend telling a “sob story” to get into law school, it is important for the committee to understand what you’ve overcome on your way here. Emphasis on the ‘overcome.’ Committees like to see how you’ve grown and changed in response to adversity. But they also want to know that those hardships are behind you and that you’ll be able to focus on your legal studies."
  },
  {
    "name": "Impressive Experiences",
    "polarity": "positive",
    "description": "We’re impressed by your résumé! At some schools, there are specific AO-codes left on files to mark someone whose experience gives them extra consideration. With the right résumé, you can punch above your weight — particularly when you’re a splitter or very close to meeting both medians."
  },
  {
    "name": "Moving Story",
    "polarity": "positive",
    "description": "We were moved by the story you told. Admissions officers are humans—while their job is to figure out why you want to come to law school (and if you’ll be a good fit), they’re going to advocate for those they feel a connection with. Telling a moving story makes it more likely for an admissions officer to fight for you—that can make a lot of difference, but they’ll need other things to point to that make you admittable."
  },
  {
    "name": "Good Writer",
    "polarity": "positive",
    "description": "We think you’re the kind of writer law schools look for. Dexterous writing is one of the best indicators of dexterous thinking. Good personal statement writing is intimate but bright, intellectual but never stuffy. Good writing is an amplifier—it doesn’t get you into a school by itself, but it does make the other elements of your application shine."
  },
  {
    "name": "Good Community Member",
    "polarity": "positive",
    "description": "We can tell that you’ve contributed significantly to past academic or professional communities. This is something admissions officers look for to estimate how you’ll contribute to your law school community. It might be useful for you to spell out how you plan to contribute to law school organizations or discussions to add a greater impact to this perspective."
  },
  {
    "name": "Lovely PS",
    "polarity": "positive",
    "description": "We loved your personal statement. That probably means you’ve both answered ‘Why Law?’ to our complete satisfaction while also using the space to tell a story that interested us. The personal statement is the place we go to ‘figure out’ as an applicant. The questions about your motivation and personality we have from your resume need to be answered by the rest of your application documents. You’ve done well."
  },
  {
    "name": "Lovely DS",
    "polarity": "positive",
    "description": "We love your diversity/perspective statement. You can think about the DS this way. Admissions officers know if you’re ‘admittable’ by the end of your personal statement. They look to the diversity statement to help them make a decision between admittable candidates. What else do you offer? The easiest way to have a great diversity statement is to offer an obviously diverse background. The other way is to evidence a strong voice and a unique way of reflecting on your life."
  },
  {
    "name": "Strong Why X",
    "polarity": "positive",
    "description": "We think you have a compelling reason for attending this particular school. Strong Why X answers are specific to both the school and the applicant. You’ve connected something about yourself to something unique that the school offers, whether that’s an academic program or just a personal connection. Some schools don’t care about Why X answers—Harvard knows everyone wants to go to Harvard. Other schools, particularly those in locations people aren’t often excited to move to, scrutinize Why X answers for evidence that you will actually come to their school. If a school doesn’t think you’ll accept their offer, they’d prefer not to make it."
  },
  {
    "name": "Lovely Supplemental",
    "polarity": "positive",
    "description": "We loved one of your supplemental statements. Supplemental essays are often places to complicate your file—to show an element of yourself that doesn’t come across in your other writing. Some, like the Yale 250, are places to show your intellectual chops. Others, like Georgetown’s Top 10 List, are simply places to demonstrate you have a winning personality."
  },
  {
    "name": "Useful addendum",
    "polarity": "positive",
    "description": "We got useful context into your situation from an addendum. Addenda should do just that: provide context we wouldn’t otherwise have. It’s not the place to tell a long story or show off your writing chops. Good addenda are simple, frank, and as short as possible. Be contrite and never defensive. If you’re explaining past behavior, show what’s changed."
  },
  {
    "name": "Rigorous Undergrad",
    "polarity": "positive",
    "description": "We can tell that you challenged yourself in undergrad beyond what your raw GPA demonstrates. While your raw UGPA is essential for determining if you help a school reach its medians, your transcript is also a story—and one that admissions officers are trained to read well. Did you challenge yourself with difficult classes? Did you take too many pass/fails without explanation? Was your major hard? Does your undergrad grade harshly compared to its peers? Did you do well despite many other commitments? These are the questions committees ask to figure out how you’ll do in the significantly more rigorous environment of a top law school."
  },
  {
    "name": "Unnecessary Addendum",
    "polarity": "negative",
    "description": "You might not need an addendum you submitted. We only recommend submitting addenda when they genuinely help add context to your file. While no two admissions readers have the exact same opinion on when an addendum is truly needed, there are some clear-cut cases. In the middle, it’s a strategic decision. Consider whether you’re actually telling the committee something they need to know. Academic addenda should provide blow-by-blow context for what went wrong and what changed in college. LSAT addenda are less useful unless they’re explaining test-day difficulties or guarding against an easy misunderstanding."
  },
  {
    "name": "Tone Check",
    "polarity": "negative",
    "description": "Even though you don’t mean it, something about your tone struck us the wrong way. The most common error in tone is inadvertently  coming across as arrogant or presumptuous. if you’re applying to law school, you want to show off your accomplishments. But sometimes things come across differently in writing. Avoid name-dropping or anything that can be misread as arrogance. This is one of the most common AO pet peeves."
  },
  {
    "name": "Questionable Formatting",
    "polarity": "negative",
    "description": "I think you need to fix the formatting of one or more documents. AOs love things to be formatted in clean, predictable ways. Your creativity should come across in your essays, not your choice of font or placement of headers. Luckily, there are usually very clear expectations when it comes to both resumes and application essays. For resumes, check out this portion of the 7Sage admissions course. For essays, read this."
  },
  {
    "name": "Generic Why X",
    "polarity": "negative",
    "description": "I think your explanation for why you want to attend a specific school is lacking. Bad Why X answers often involve things that could be true about any school. Cornell knows that Columbia also has clinics and experiential learning opportunities that will allow you to dive into immigration law. What it doesn’t know is if you’d prefer the Finger Lakes to the Upper West Side. If you’re going to focus on academic reasons for preferring one school over the other, make sure that they’re legitimately unique."
  },
  {
    "name": "Test History Needs Explanation",
    "polarity": "negative",
    "description": "Your LSAT or GRE history left us wanting context. This can be a tricky one. Some admissions readers want an explanation for as little as a six-point jump. Often, the answer is just a simple ‘I studied more.’ Check school instructions for specific guidance. Generally, if you’ve canceled your score twice or have a very significant outlier, an explanation doesn’t hurt. If your score history reads 160, 162, 158, 172, C, almost any reader will be desperate to know why they should think that 172 is actually more representative of your abilities."
  },
  {
    "name": "Academics Need Explanation",
    "polarity": "negative",
    "description": "Looking over your transcript and/or UPGA, we need context. Schools expect academic addenda when you had one or two semesters that were significantly worse without apparent explanation or if you’ve taken too many withdrawals or pass/fails. Avoid writing an academic addendum for a single C. But multiple Cs or a D or F likely warrants an explanation."
  },
  {
    "name": "Résumé Rehash",
    "polarity": "negative",
    "description": "You’re spending too much time repeating your accomplishments in your essays. AOs often look for more information about your experiences in your personal statement. Something becomes a “résumé rehash” when it feels like you’re going through too many experiences and mostly focusing on the surface level. Moving chronologically “I did this, then I did this” is the biggest sign of a résumé rehash. You should probably revise your personal statement in one of two ways. Either pick one or two experiences to go dive into with much more detail — or make sure that there’s a clear throughline through each paragraph that isn’t just about what you did next: what ideas, questions, and personal quests drove your movement through your various experiences?"
  },
  {
    "name": "Résumé Needs Polish",
    "polarity": "negative",
    "description": "Your résumé could use more polish. When it comes to résumés, consistency is the most important aspect. Make sure that your formatting for each section is clear and identical—down to consistently using the same kind of dash between dates (one of the most common mistakes.) The best place to start is with a clear, law-school approved template. We recommend one of our own."
  },
  {
    "name": "Padded Résumé",
    "polarity": "negative",
    "description": "It seemed like you were trying to fluff up your résumé. If you’re onto a second page, consider going down to just one page. Don’t try to overplay your hand: three bullets are almost always enough even for complex roles. Don’t try to play limited volunteer engagements off as full employment."
  },
  {
    "name": "Cramped Résumé",
    "polarity": "negative",
    "description": "There wasn’t enough white space in your résumé— it made it difficult to read. AOs need at least .75” margins on all sides — and most prefer 1-inch. Readability is the key — your résumé is essential, but AOs aren’t spending too much time here. Readability depends on two things: information being where AOs expect it (dates on the right, education at the top, personal at the bottom) and breathability. The whitespace on your résumé is essential. If you haven’t gone onto a second page, consider it. Also consider if you can cut back on some bullets (three is enough for almost every position.)"
  },
  {
    "name": "Repetitive Statements",
    "polarity": "negative",
    "description": "You’re repeating too much material between different parts of your application. An application isn’t like a college paper — you actually don’t need to fill up every space. Only write an essay if it will add a significant new element to your file. Often applicants struggle to avoid repeating material between their personal statement and ‘perspective’ statement. Put your most important story, the one that really explains who you are and why you’re going to law school, in the PS. Then use your diversity statement to show some other element of who you are."
  },
  {
    "name": "A Little Stiff",
    "polarity": "negative",
    "description": "Your personality isn’t coming through strongly enough in your writing. After we read files, AOs ask “Do I feel like I really know this person?” Sounding too formal sends the wrong signal. Yes, schools need to know that you can be professional. But they also need to know that you’re a human with quirks and passions. Imagine writing a letter to a professor you’re very comfortable with. You want to still present yourself at your best, of course, but you also want to connect with them. That’s your goal here."
  },
  {
    "name": "Degree Collector?",
    "polarity": "negative",
    "description": "A lot of people apply to law school after doing previous graduate work. A master’s can be a great way to show academic potential, personal growth, and commitment to your interests. But if you’re applying to law school after multiple MAs, or after an MBA or PhD, committees might start to wonder if you’re more interested in being a law student than a lawyer. Because law school rankings depend on job outcomes, schools are fastidious in filtering out those whose J.D.-motivations they don’t understand. The antidote to this is a strong Why Law — spend more of your application explaining, and then proving, why you really do need a J.D., and why you’re certain this professional path is now the only one for you. If a J.D. can be convincingly portrayed as a continuation of your studies, turning your previous graduate studies into a compelling advantage you’ll have, make that argument. If it’s a genuine change of direction, tell a story that gives your reader insight."
  },
  {
    "name": "Be More Professional",
    "polarity": "negative",
    "description": "While it’s important not to come across as too stiff, something in your file made us question your professionalism. We mean no offense: law school sits on the intersection of academia and the professional world. 1Ls often come in little more than college students. They leave prepared to make decisions with the highest human stakes. We look to the application to figure out if you can walk a fine line between the personal and the professional. Humor is good, but don’t be too jokey. You can be open about things you’ve gone through, but be mindful of reader’s sensibilities. And don’t presume you know the law before you’ve become a lawyer—committees often respond negatively to that."
  },
  {
    "name": "C&F Not Sufficiently Addressed",
    "polarity": "negative",
    "description": "You needed to write an explanation for a character & fitness issue, and it’s still not quite there. We still have fundamental questions about what happened or how you’ve changed as a result. C&F addenda are tricky—while of course you shouldn’t reveal anything that you don’t have to, you do need to be open and contrite in response to a university’s questions. Avoid anything that seems like you’re making an excuse for yourself. Give context in a clear, contrite tone."
  },
  {
    "name": "Show More Maturity",
    "polarity": "negative",
    "description": "A lot of people apply to law school when they’re still very young. For KJDs or other applicants without extensive work experiences, law school admissions officers need to find the ones who seem like they’re already ready for the working world and civil but heated classroom discussions with people who might profoundly disagree with you. Don’t force it, but try to spend more of your application demonstrating a mature personality and reflections. If you’re particularly young, avoid portraying yourself as young in your personal statement by focusing too much on early life."
  },
  {
    "name": "Writing Needs Improvement",
    "polarity": "negative",
    "description": "Law school admissions officers look at your writing as evidence of your thinking. It’s very, very important that you’re writing at your best. And we do mean your best. Law schools have your LSAT writing on file and will notice if your writing voice differs drastically between the two documents. That doesn’t mean that you shouldn’t put your writing through multiple rounds of revision. Focus specifically on a clear sequence of ideas, paragraphs that each serve a unique purpose, and prose that’s easy-to-follow but still vivid an unique."
  },
  {
    "name": "Show More Openness",
    "polarity": "negative",
    "description": "Strong opinions are good, but law schools look for extra reassurance that you can “connect across differences.” Law schools are more ideologically diverse than most undergrad classrooms — there will be students (and professors) with radically different opinions than you. Don’t hide key parts of who you are, but do signal that you’re excited to connect with people who disagree with you through productive discussion."
  },
  {
    "name": "Unexplained Gap",
    "polarity": "negative",
    "description": "You have a gap in your work history or academic record I need an explanation for. Sometimes it’s possible to quickly address a gap on your résumé by showing what you were doing during that time, such as taking care of a family member. If you took academic leave or have been out of work for a while, committees will likely expect an addendum to explain why you took that time. As always, they’re worried about students who might take similar gaps after law school, which would hurt their employment rankings. Don’t overdo it—just give the needed context."
  },
  {
    "name": "Unforced Error",
    "polarity": "negative",
    "description": "While they want you to be contrite, law schools also don’t want you to disclose unnecessarily harmful information. If you did something against the rules but were never caught, now isn’t the time for a confession. Law schools look for the lawyerly skill of knowing what to reveal when. Make sure you know how to show your best face."
  },
  {
    "name": "Too Much Early Life",
    "polarity": "negative",
    "description": "Reading your file, we learned too much about your childhood and not enough about who you are now. Now isn’t the time to tackle a topic that could have worked as a college application essay. Introduce yourself to AOs as a fully-formed adult—then go back to establish things about your early life if they’re necessary to understanding why you’re seeking a law degree today."
  },
  {
    "name": "AI?",
    "polarity": "negative",
    "description": "Even if you didn’t use it, something in your statements sounded like AI. Our admissions readers are trained to detect AI-generated text. If an AO suspects you composed your application documents with generative AI, the safest thing they can do is auto-deny. The ethical demands of the legal profession are profound and if they can’t trust you to apply without AI, they can’t trust you to actually do the work required in law school. If you used AI to write your statements, rewrite them without AI. If this is a false positive, you should identify the features of your writing giving that impression and revise to err on the side of caution."
  },
  {
    "name": "Workplace-Bound",
    "polarity": "negative",
    "description": "Legacy tag from committee ballots. Explanation not yet provided in the tag explanation CSV."
  },
  {
    "name": "Stretched Résumé",
    "polarity": "negative",
    "description": "Legacy tag from committee ballots. Explanation not yet provided in the tag explanation CSV."
  },
  {
    "name": "Unnecessary Disclosure",
    "polarity": "negative",
    "description": "Legacy tag from committee ballots. Explanation not yet provided in the tag explanation CSV."
  }
];

const TAG_NAME_SET = new Set(TAG_DEFINITIONS.map((tag) => tag.name));
const TAG_POLARITY_MAP = new Map(TAG_DEFINITIONS.map((tag) => [tag.name, tag.polarity]));
const TAG_DESCRIPTION_MAP = new Map(TAG_DEFINITIONS.map((tag) => [tag.name, tag.description]));
const TAG_CANONICAL_NAME_MAP = new Map(
  TAG_DEFINITIONS.map((tag) => [normalizeTagKey(tag.name), tag.name])
);
const HIDDEN_TAG_NAMES = new Set([
  "Cramped Résumé",
  "Workplace-Bound",
  "Stretched Résumé",
  "Unnecessary Disclosure",
]);
const DISPLAY_TAG_DEFINITIONS = TAG_DEFINITIONS.filter(
  (tag) => !HIDDEN_TAG_NAMES.has(tag.name)
);
const READER_PROFILES = [
  {
    fullName: "Tajira McCoy",
    firstName: "Tajira",
    aliases: ["Tajira McCoy", "Tajira"],
    headshotUrl:
      "https://www.gravatar.com/avatar/9cefff863fcd2cd8f5301fe9a68caf0b?size=320&default=robohash",
    bio:
      "Former Director of Admissions & Scholarships at Berkeley Law. 7Sage expert in scholarships & diversity.",
  },
  {
    fullName: "Brigitte Suhr",
    firstName: "Brigitte",
    aliases: ["Brigitte", "Briggite"],
    headshotUrl:
      "https://www.gravatar.com/avatar/759b37b8091b593596f86f07072fe396?size=320&default=robohash",
    bio:
      "Former UVA AO. Accomplished international human rights lawyer.",
  },
  {
    fullName: "Dr. Sam Riley",
    firstName: "Sam",
    aliases: ["Dr. Sam Riley", "Sam Riley"],
    headshotUrl:
      "https://www.gravatar.com/avatar/900993889789621e14e63ad6a9958ca7?size=320&default=robohash",
    bio:
      "Senior Director of Admissions Programs at Texas Law. 7Sage's expert on resume & KJD admissions.",
  },
  {
    fullName: "Sam Kwak",
    firstName: "Sam",
    aliases: ["Sam Kwak"],
    headshotUrl:
      "https://www.gravatar.com/avatar/c47ae9e0cd28ca0ad23e082ae6f6df1f?size=320&default=robohash",
    bio:
      "Former Stanford Law AO. Former Senior Associate Director of Admissions at Northwestern Law.",
  },
  {
    fullName: "Reyes Aguilar",
    firstName: "Reyes",
    aliases: ["Reyes Aguilar", "Reyes"],
    headshotUrl:
      "https://www.gravatar.com/avatar/5df779a3d2aef8fbdd29088b8cc485d2?size=320&default=robohash",
    bio:
      "30 years of experience in law school admissions at Utah Law. Currently reviews files at UCLA Law.",
  },
  {
    fullName: "Kory Hawkins",
    firstName: "Kory",
    aliases: ["Kory Hawkins", "Kory"],
    headshotUrl:
      "https://www.gravatar.com/avatar/74291fcb1e0ab570fd7e026e57dd278f?size=320&default=robohash",
    bio:
      "Former Columbia Law AO. Former Associate Director of Admissions at UC Law SF.",
  },
  {
    fullName: "Kamil Brown",
    firstName: "Kamil",
    aliases: ["Kamil Brown", "Kamil"],
    headshotUrl:
      "https://www.gravatar.com/avatar/f3e721c5dc9708ebc13125ecd0db7716?size=320&default=robohash",
    bio:
      "Former Yale Law AO. Former admissions reader at Fordham Law.",
  },
  {
    fullName: "Jacob Baska",
    firstName: "Jacob",
    aliases: ["Jacob Baska", "Jacob", "Jake Baska", "Jake"],
    headshotUrl:
      "https://www.gravatar.com/avatar/078822ac5e3142951af136795491c23c?size=320&default=robohash",
    bio:
      "Former Director of Admissions at Notre Dame Law. Host of the 7Sage Admissions Podcast.",
  },
  {
    fullName: "Jocelyn Glantz",
    firstName: "Jocelyn",
    aliases: ["Jocelyn Glantz", "Jocelyn"],
    headshotUrl:
      "https://www.gravatar.com/avatar/47947e4483a165443bbb98abecfdde40?size=320&default=robohash",
    bio: "Former Assistant Director of Admissions at Brooklyn Law.",
  },
  {
    fullName: "Jen Kott",
    firstName: "Jen",
    aliases: ["Jen Kott", "Jennifer Kott", "Jen", "Jennifer"],
    headshotUrl:
      "https://www.gravatar.com/avatar/29d0502c4ea11721ff29d3d1fa1c3bdd?size=320&default=robohash",
    bio:
      "Former AO at Northwestern, Tulane, & UNC Chapel Hill. Former Director of Admissions at Arizona Law.",
  },
  {
    fullName: "Bette Bradley",
    firstName: "Bette",
    aliases: ["Bette Bradley", "Bette"],
    headshotUrl:
      "https://www.gravatar.com/avatar/c16339392d0ad8e50f41bb057d7a17d3?size=320&default=robohash",
    bio:
      "Former Assistant Dean for Admissions & Scholarships at Ole Miss Law. Accomplished healthcare attorney.",
  },
  {
    fullName: "Sherolyn Hurst",
    firstName: "Sherolyn",
    aliases: ["Sherolyn Hurst", "Sherolyn"],
    headshotUrl:
      "https://www.gravatar.com/avatar/66552b1ad0715e56128d3163639a4ba4?size=320&default=robohash",
    bio:
      "Former Director of Admissions at Golden Gate Law. Former AO at Dedman, St. Thomas, & Texas Wesleyan.",
  },
  {
    fullName: "Carol Cochran",
    firstName: "Carol",
    aliases: ["Carol Cochran", "Carol"],
    headshotUrl:
      "https://www.gravatar.com/avatar/cd528c47bc4c30f6418a7f422a9407b1?size=320&default=robohash",
    bio: "Assistant Dean at Seattle University School of Law.",
  },
];

const TAG_FONT_BASE_PX = 9;
const TAG_FONT_MIN_PX = 7;

const state = {
  fileName: "",
  parsedRows: [],
  groupedRows: new Map(),
  availableStudents: [],
  studentInputByFile: {},
  reports: [],
  currentReport: null,
  warnings: [],
  errors: [],
};

let fitResizeScheduled = false;

const PRINT_CSS = `
  @import url("https://fonts.googleapis.com/css2?family=Fraunces:wght@600;700&family=Lexend:wght@400;500;600;700;800&display=swap");
  *, *::before, *::after { box-sizing: border-box; }
  @page { size: 8.5in 11in; margin: 0; }
  body { margin: 0; font-family: "Lexend", "Segoe UI", Tahoma, sans-serif; color: #202530; background: #fffffe; }
  .doc-shell { background: #fff; }
  .page { width: 612px; min-height: 792px; height: 792px; padding: 24px; margin: 0; break-after: page; background: #fffffe; }
  .page.summary-page { padding: 0; min-height: 792px; height: 792px; width: 612px; transform: scale(1.3333333333); transform-origin: top left; margin-bottom: 264px; background: #fffffe; }
  .page:last-child { break-after: auto; }
  .page-header { border-bottom: 2px solid #111; padding-bottom: 0.1in; margin-bottom: 0.2in; }
  .doc-title { margin: 0; font-size: 20pt; }
  .subtitle { margin: 4px 0 0; color: #5e6778; font-size: 11pt; }
  .row { display: grid; grid-template-columns: 180px 1fr; gap: 10px; margin-bottom: 10px; align-items: baseline; }
  .label { font-weight: 600; }
  .value { min-height: 1.2em; }
  .small { font-size: 10pt; color: #5e6778; }
  .rating-row { border: 1px solid #d8dee9; border-radius: 8px; padding: 8px 10px; margin-bottom: 8px; }
  .stars-wrap { display: flex; align-items: center; gap: 8px; margin-top: 6px; }
  .stars { display: inline-flex; gap: 4px; }
  .star { width: 18px; height: 18px; }
  .star.filled polygon { fill: #f4b400; stroke: #b78600; }
  .star.half polygon { stroke: #b78600; }
  .waiting-page { display: flex; align-items: center; justify-content: center; }
  .waiting-copy { font-family: "Fraunces", "Times New Roman", serif; font-size: 36px; font-weight: 700; color: #98a2b3; text-transform: lowercase; }
  .band-list { margin: 0; padding-left: 18px; }
  .band-list li { margin-bottom: 6px; }
  .tag-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
  .tag-pill { --tag-font-size: 9px; border: 1px solid #e0e6f2; border-radius: 999px; padding: 2px 6px; font-size: var(--tag-font-size); line-height: 1.1; font-weight: 600; text-align: center; display: flex; align-items: center; justify-content: center; position: relative; overflow: visible; color: #4b5563; }
  .tag-text { display: -webkit-box; width: 100%; overflow: hidden; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
  .tag-pill.active-positive { background: #ecfdf3; border-color: #86efac; color: #166534; }
  .tag-pill.active-negative { background: #fef2f2; border-color: #fecaca; color: #7f1d1d; }
  .tag-badges { display: inline-flex; gap: 4px; position: absolute; top: -12px; right: 8px; margin-left: 0; vertical-align: baseline; z-index: 2; }
  .tag-badge { min-width: 14px; height: 14px; border-radius: 999px; display: inline-flex; align-items: center; justify-content: center; font-size: 9px; font-weight: 700; background: #111827; color: #fff; border: 1px solid #fff; }
  .tag-badge.reader-slot-1 { background: #227f9c; }
  .tag-badge.reader-slot-2 { background: #15b79e; }
  .tag-badge.reader-slot-3 { background: #db2777; }
  .tag-pill.inactive { color: #5e6778; background: #fff; }
  .reader-cards-source { display: none; }
  .reader-cards { display: flex; flex-direction: column; gap: 12px; height: 100%; flex: 1; }
  .reader-card { border: 2px solid #d8dee9; border-radius: 12px; padding: 12px; display: grid; gap: 10px; }
  .reader-card.reader-slot-1 { border-color: #227f9c; }
  .reader-card.reader-slot-2 { border-color: #15b79e; }
  .reader-card.reader-slot-3 { border-color: #db2777; }
  .page-continue-note { position: absolute; right: 18px; bottom: 12px; font-size: 9px; letter-spacing: 0.08em; text-transform: uppercase; color: #94a3b8; opacity: 0; }
  .page.has-continue .page-continue-note { opacity: 1; }
  .reader-title { margin: 0 0 8px; font-size: 14px; }
  .notes-box { border: 1px solid #333; min-height: 1.5in; padding: 10px; white-space: pre-wrap; margin-bottom: 8px; }
  .summary-banner { height: 58px; background: linear-gradient(-45deg, #15b79e 0%, #227f9c 100%); color: #fcfaf8; text-align: center; font-family: "Fraunces", "Times New Roman", serif; font-size: 30px; font-weight: 700; line-height: 58px; margin-bottom: 0; }
  .section-block { display: grid; grid-template-columns: 1fr; gap: 0; margin-bottom: 0; width: 100%; background: #fff; border-top: 0; border-bottom: 0; }
  .summary-page .section-block { background: #fffffe; }
  .section-block.basics-section { height: 52px; }
  .section-block.readers-section { height: 150px; margin-top: 0; }
  .section-block.readers-section .section-body { height: 100%; overflow: hidden; padding: 2px 8px; }
  .section-block.readers-section .avatar { width: 36px; height: 36px; border-radius: 9px; font-size: 11px; }
  .section-block.readers-section .reader-name { font-size: 13px; }
  .section-block.readers-section .reader-bio { font-size: 10px; line-height: 1.2; -webkit-line-clamp: 4; }
  .section-block.readers-section .reader-bio-list { font-size: 10px; line-height: 1.2; margin: 0; padding-left: 14px; display: grid; gap: 2px; max-height: calc(4 * 1.2em + 6px); overflow: hidden; }
  .section-block.readers-section .reader-bio-list li { margin: 0; }
  .fit-readers { display: flex; flex-direction: column; height: 100%; }
  .readers-card { border: 2px solid #d8dee9; border-radius: 12px; padding: 4px 6px; flex: 1; min-height: 0; }
  .section-block.basics-section .section-body { height: 100%; padding: 4px 10px; }
  .section-block.key-takeaways-section { height: 232px; margin-top: 0; }
  .section-block.key-takeaways-section .section-body { height: 100%; padding: 3px 12px; }
  .section-block.tags-section { height: 300px; border-bottom: 0; }
  .section-block.tags-section .section-body { height: 100%; overflow: hidden; }
  .section-block.tags-section .section-body { padding: 6px 8px; }
  .fit-tags { display: flex; flex-direction: column; height: 100%; }
  .tags-card { border: 2px solid #d8dee9; border-radius: 12px; padding: 6px 8px; flex: 1; display: flex; flex-direction: column; }
  .tags-grid-wrap { flex: 1; display: flex; align-items: center; justify-content: center; min-height: 0; width: 100%; }
  .fit-tags .design-tag-grid { transform: scale(0.93); transform-origin: center; width: 100%; height: 100%; }
  .basics-grid { display: flex; align-items: center; justify-content: center; gap: 10px; flex-wrap: nowrap; }
  .basics-item { display: inline-flex; gap: 4px; align-items: baseline; text-align: left; }
  .basics-label { font-size: 10px; font-weight: 700; letter-spacing: 0.12em; text-transform: uppercase; color: #475467; }
  .basics-value { font-size: 12px; font-weight: 700; }
  .basics-divider { font-size: 12px; color: #98a2b3; padding: 0 4px; }
  .takeaways-row { display: grid; gap: 8px; margin-bottom: 8px; }
  .takeaways-row-top { grid-template-columns: repeat(3, minmax(0, 1fr)); }
  .takeaways-row-bottom { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .takeaway-item { text-align: center; display: grid; gap: 4px; }
  .takeaway-title { font-family: "Fraunces", "Times New Roman", serif; font-size: 10px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.1em; color: #475467; }
  .takeaway-softs-value { font-size: 14px; font-weight: 700; }
  .takeaways-bands .band-row { font-size: 9px; gap: 0; padding: 3px 6px; grid-template-columns: 52px repeat(9, minmax(0, 1fr)); }
  .takeaways-card { border: 2px solid #d8dee9; border-radius: 12px; padding: 8px 10px; display: grid; gap: 8px; flex: 1; min-height: 0; }
  .fit-takeaways { display: flex; flex-direction: column; height: 100%; }
  .section-title { font-family: "Fraunces", "Times New Roman", serif; text-transform: uppercase; letter-spacing: 0.22em; font-size: 10px; font-weight: 700; color: #94a3b8; text-align: center; margin: 0 0 4px; }
  .section-block.readers-section .section-title { margin-bottom: 1px; }
  .rail-label { writing-mode: vertical-rl; transform: rotate(180deg); background: #334155; color: #fff; width: 40px; border-radius: 18px 0 0 18px; text-align: center; font-family: "Fraunces", "Times New Roman", serif; font-size: 22px; font-weight: 700; letter-spacing: 0.4px; padding: 16px 8px; }
  .section-body { border: 0; border-radius: 0; padding: 14px 16px; background: transparent; width: 100%; min-width: 0; overflow: hidden; }
  .fit-content { transform-origin: top left; width: 100%; }
  .reader-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); grid-template-rows: 1fr; gap: 10px; height: 100%; align-content: stretch; }
  .reader-col { padding: 2px 0; position: relative; display: flex; flex-direction: column; align-items: center; gap: 4px; min-height: 0; }
  .summary-page .reader-col:not(:last-child)::after { content: ""; position: absolute; top: 10%; bottom: 10%; right: -6px; width: 2px; background: #d8dee9; }
  .avatar { width: 52px; height: 52px; border-radius: 12px; margin: 0; display: flex; align-items: center; justify-content: center; background: linear-gradient(135deg, #2b7abf, #14b8a6); color: #fff; font-weight: 800; font-size: 14px; border: 1px solid transparent; overflow: hidden; }
  .avatar.has-photo { background: transparent; padding: 0; }
  .avatar.reader-slot-1 { border-color: #227f9c; }
  .avatar.reader-slot-2 { border-color: #15b79e; }
  .avatar.reader-slot-3 { border-color: #db2777; }
  .avatar img { width: 100%; height: 100%; border-radius: 0; object-fit: cover; display: block; }
  .reader-name { text-align: center; margin: 0 0 1px; font-size: 16px; font-family: "Fraunces", "Times New Roman", serif; white-space: nowrap; text-overflow: ellipsis; overflow: hidden; }
  .reader-bio { margin: 0; text-align: center; font-size: 11px; line-height: 1.3; display: -webkit-box; -webkit-line-clamp: 4; -webkit-box-orient: vertical; overflow: hidden; }
  .reader-bio-list { margin: 0; padding-left: 16px; text-align: left; font-size: 11px; line-height: 1.3; display: grid; gap: 2px; max-height: calc(4 * 1.3em + 6px); overflow: hidden; }
  .reader-bio-list li { margin: 0; }
  .key-card { border: 0; border-radius: 0; overflow: hidden; width: 100%; height: 100%; display: flex; flex-direction: column; }
  .key-top { display: grid; grid-template-columns: minmax(0, 0.8fr) minmax(0, 0.8fr) minmax(0, 1.4fr) minmax(0, 0.9fr); border-bottom: 2px solid #d8dee9; }
  .key-top-item { padding: 8px 12px; font-size: 19px; font-weight: 700; white-space: nowrap; position: relative; }
  .key-top-label { opacity: 0.9; }
  .key-top-value { margin-left: 4px; font-weight: 800; }
  .key-top-item.other-item .key-top-value { font-size: 13px; font-weight: 500; letter-spacing: 0.1px; }
  .key-top-item.softs-item .key-top-value { font-weight: 500; }
  .key-top-item:not(:last-child)::after { content: ""; position: absolute; top: 50%; right: 0; transform: translateY(-50%); width: 2px; height: 60%; background: #d8dee9; }
  .metric-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); border-bottom: 2px solid #d8dee9; }
  .metric-col { padding: 8px 10px; text-align: center; position: relative; }
  .metric-col:not(:last-child)::after { content: ""; position: absolute; top: 50%; right: 0; transform: translateY(-50%); width: 2px; height: 60%; background: #d8dee9; }
  .metric-title { margin: 0 0 6px; font-size: 20px; font-weight: 700; }
  .compact-stars { display: inline-flex; gap: 3px; justify-content: center; }
  .compact-stars .star { width: 14px; height: 14px; }
  .band-row { display: grid; grid-template-columns: 72px repeat(9, minmax(0, 1fr)); align-items: center; gap: 0; padding: 5px 8px; font-size: 11px; position: relative; flex: 1 0 0; }
  .band-name { font-weight: 700; border-radius: 999px; text-align: center; padding: 2px 6px; margin-right: 6px; position: relative; z-index: 3; }
  .band-name.reach { background: #fff2df; color: #b45309; }
  .band-name.target { background: #e0f2fe; color: #0c4a6e; }
  .band-name.safety { background: #dcfce7; color: #14532d; }
  .band-row.row-reach .band-chip.in-range { color: #b45309; }
  .band-row.row-target .band-chip.in-range { color: #0c4a6e; }
  .band-row.row-safety .band-chip.in-range { color: #14532d; }
  .band-range { grid-row: 1; border-radius: 8px; min-height: 18px; align-self: stretch; z-index: 1; }
  .band-range.range-reach { background: #fff2df; border: 1px solid #f3ce73; }
  .band-range.range-target { background: #e0f2fe; border: 1px solid #7dd3fc; }
  .band-range.range-safety { background: #dcfce7; border: 1px solid #86efac; }
  .band-chip { display: flex; align-items: center; justify-content: center; text-align: center; padding: 2px 0; min-height: 18px; font-weight: 700; color: #111827; position: relative; z-index: 2; }
  .design-tag-grid { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); grid-template-rows: repeat(7, minmax(0, 1fr)); gap: 8px; height: 100%; align-content: stretch; flex: 1; min-height: 0; }
  .page.reader-detail-page { padding: 18px; display: flex; flex-direction: column; gap: 10px; background: #fffffe; position: relative; }
  .reader-top-grid { display: grid; grid-template-columns: 1.35fr 0.6fr 1.05fr; gap: 10px; min-height: 0; }
  .page.reader-summary-page { padding: 24px; background: #fffffe; }
  .reader-summary-grid { display: grid; grid-template-columns: 1fr; grid-template-rows: repeat(2, minmax(0, 1fr)); gap: 12px; height: 100%; }
  .reader-footer-box { border: 1px solid #d8dee9; border-radius: 8px; padding: 6px 8px; display: grid; grid-template-rows: auto 1fr; gap: 4px; overflow: hidden; font-size: 10px; line-height: 1.3; }
  .reader-footer-title { font-family: "Fraunces", "Times New Roman", serif; font-weight: 700; font-size: 9px; text-transform: uppercase; letter-spacing: 0.12em; color: #475467; }
  .reader-footer-body { white-space: pre-wrap; }
  .reader-section-title { font-family: "Fraunces", "Times New Roman", serif; font-weight: 700; font-size: 10px; letter-spacing: 0.18em; text-transform: uppercase; color: #344054; margin-bottom: 0; text-align: center; }
  .reader-detail-page .reader-col { padding: 0; position: relative; display: grid; grid-template-columns: 1fr; align-items: start; }
  .reader-col.ratings, .reader-col.bands { display: grid; grid-template-columns: 1fr; grid-auto-rows: minmax(0, 1fr); gap: 6px; min-height: 0; }
  .reader-col.tags { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 6px; align-content: start; min-height: 0; }
  .reader-rating-row { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
  .reader-rating-label { font-size: 8px; font-weight: 600; text-transform: none; letter-spacing: 0; color: #475467; min-width: 0; flex: 1; line-height: 1.2; }
  .reader-rating-pill { border: 1px solid #e2e8f0; border-radius: 999px; padding: 2px 6px; font-size: 8px; font-weight: 600; line-height: 1.1; background: #f8fafc; color: #334155; white-space: nowrap; }
  .reader-rating-pill.is-positive, .reader-rating-pill.is-positive-strong { background: #ecfdf3; border-color: #34d399; color: #166534; }
  .reader-rating-pill.is-negative, .reader-rating-pill.is-negative-strong { background: #fef2f2; border-color: #f87171; color: #991b1b; }
  .reader-rating-pill.is-positive, .reader-rating-pill.is-negative { opacity: 0.5; }
  .reader-rating-pill.empty { color: #94a3b8; background: #fff; }
  .reader-rating-empty { font-size: 8px; color: #98a2b3; }
  .reader-band-row { display: grid; grid-template-columns: auto 1fr; gap: 4px; align-items: center; font-size: 9px; }
  .reader-band-label { font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; font-size: 8px; color: #475467; }
  .reader-band-value { font-weight: 600; }
  .reader-notes { border: 0; border-radius: 0; padding: 0; font-size: 11px; line-height: 1.35; height: 100%; overflow: hidden; }
  .reader-notes .label { font-weight: 700; font-size: 10px; text-transform: uppercase; letter-spacing: 0.12em; margin-bottom: 2px; color: #475467; }
  .reader-notes-body { margin: 0; white-space: pre-wrap; }
  .reader-tag-pill { border: 1px solid #e0e6f2; border-radius: 999px; padding: 1px 3px; font-size: 7px; line-height: 1.1; text-align: center; display: flex; align-items: center; justify-content: center; min-height: 16px; }
  .reader-tag-pill.active-positive { background: #ecfdf3; border-color: #86efac; color: #166534; }
  .reader-tag-pill.active-negative { background: #fef2f2; border-color: #fecaca; color: #7f1d1d; }
  .reader-tag-pill.inactive { color: #5e6778; background: #fff; }
  .reader-tag-empty { font-size: 10px; color: #98a2b3; align-self: center; }
  .tag-explanation-source { display: none; }
  .tag-explanation-page { padding: 24px; background: #fffffe; display: flex; flex-direction: column; height: 792px; position: relative; }
  .tag-explanation-cards { display: flex; flex-direction: column; gap: 12px; flex: 1; min-height: 0; padding-bottom: 22px; }
  .tag-explanation-item { border: 1px solid #d8dee9; border-radius: 8px; padding: 10px 12px; background: transparent; display: grid; gap: 6px; min-height: 0; flex: 0 0 auto; }
  .tag-explanation-title { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
  .tag-explanation-pill { padding: 4px 10px; font-size: 10px; line-height: 1.1; }
  .tag-explanation-cont { font-size: 10px; line-height: 1; color: #98a2b3; font-weight: 600; }
  .tag-explanation-body { font-size: 11px; line-height: 1.4; color: #1f2937; }
  .tag-explanation-empty { display: flex; align-items: center; justify-content: center; height: 100%; font-size: 14px; color: #98a2b3; }
`;

const PRINT_FIT_SCRIPT = `
  (() => {
    const TAG_FONT_BASE_PX = 9;
    const TAG_FONT_MIN_PX = 7;

    function fitTagFonts(root = document) {
      const pills = root.querySelectorAll(".tag-pill");
      pills.forEach((pill) => {
        const textEl = pill.querySelector(".tag-text");
        if (!textEl) return;

        const prevDisplay = textEl.style.display;
        const prevClamp = textEl.style.webkitLineClamp;
        const prevOrient = textEl.style.webkitBoxOrient;

        textEl.style.display = "block";
        textEl.style.webkitLineClamp = "unset";
        textEl.style.webkitBoxOrient = "initial";

        const setFont = (sizePx) => {
          pill.style.setProperty("--tag-font-size", sizePx + "px");
        };

        const getLineHeight = () => {
          const computed = window.getComputedStyle(textEl);
          const fontSize = parseFloat(computed.fontSize) || TAG_FONT_BASE_PX;
          let lineHeight = parseFloat(computed.lineHeight);
          if (!lineHeight || Number.isNaN(lineHeight)) {
            lineHeight = fontSize * 1.1;
          }
          return { lineHeight, fontSize };
        };

        setFont(TAG_FONT_BASE_PX);
        let { lineHeight } = getLineHeight();
        let maxHeight = lineHeight * 2 + 0.5;
        let fullHeight = textEl.scrollHeight;

        if (fullHeight > maxHeight) {
          let low = TAG_FONT_MIN_PX;
          let high = TAG_FONT_BASE_PX;
          let best = TAG_FONT_MIN_PX;

          while (low <= high) {
            const mid = Math.floor((low + high) / 2);
            setFont(mid);
            ({ lineHeight } = getLineHeight());
            maxHeight = lineHeight * 2 + 0.5;
            fullHeight = textEl.scrollHeight;

            if (fullHeight <= maxHeight) {
              best = mid;
              low = mid + 1;
            } else {
              high = mid - 1;
            }
          }
          setFont(best);
        }

        textEl.style.display = prevDisplay;
        textEl.style.webkitLineClamp = prevClamp;
        textEl.style.webkitBoxOrient = prevOrient;
      });
    }

    function paginateReaderCards(root = document) {
  const scope = root.querySelector(".doc-shell") || root;
  const source = scope.querySelector("[data-reader-source]");
  if (!source) return;

  scope.querySelectorAll(".reader-detail-page").forEach((page) => page.remove());

  const parent = source.parentNode;
  const anchor = source.nextSibling;
  const cardTemplates = [...source.querySelectorAll(".reader-card")];
  if (!cardTemplates.length) return;

  const createPage = () => {
    const page = document.createElement("section");
    page.className = "page reader-detail-page";
    page.innerHTML =
      '<div class="reader-cards"></div><div class="page-continue-note">Section continues on next page.</div>';
    return page;
  };

  const pages = [];
  let currentPage = createPage();
  parent.insertBefore(currentPage, anchor);
  pages.push(currentPage);
  let container = currentPage.querySelector(".reader-cards");

  const addPage = () => {
    currentPage = createPage();
    parent.insertBefore(currentPage, anchor);
    pages.push(currentPage);
    container = currentPage.querySelector(".reader-cards");
  };

  const tryAppendCard = (card) => {
    container.appendChild(card);
    const fits = container.scrollHeight <= container.clientHeight + 0.5;
    if (!fits) container.removeChild(card);
    return fits;
  };

  const splitCardIntoPages = (baseCard, fullText) => {
    let remaining = fullText.trim();
    let first = true;
    while (remaining) {
      let card = first ? baseCard : baseCard.cloneNode(true);
      if (!first) {
        card.classList.add("continued");
        const header = card.querySelector(".reader-section-title");
        if (header && !header.textContent.includes("(cont.)")) {
          header.textContent = header.textContent + " (cont.)";
        }
        const topGrid = card.querySelector(".reader-top-grid");
        if (topGrid) topGrid.remove();
      }

      const notesBody = card.querySelector(".reader-notes-body");
      if (!notesBody) {
        container.appendChild(card);
        return;
      }

      const tokens = remaining.match(/\S+|\s+/g) || [];
      container.appendChild(card);

      let low = 1;
      let high = tokens.length;
      let best = 0;

      while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        notesBody.textContent = tokens.slice(0, mid).join("");
        if (container.scrollHeight <= container.clientHeight + 0.5) {
          best = mid;
          low = mid + 1;
        } else {
          high = mid - 1;
        }
      }

      if (best <= 0) {
        notesBody.textContent = remaining;
        return;
      }

      notesBody.textContent = tokens.slice(0, best).join("");
      remaining = tokens.slice(best).join("").trimStart();

      if (remaining) {
        addPage();
        first = false;
      } else {
        return;
      }
    }
  };
  cardTemplates.forEach((template) => {
    const card = template.cloneNode(true);
    const notesBody = card.querySelector(".reader-notes-body");
    if (notesBody && !notesBody.dataset.fullText) {
      notesBody.dataset.fullText = notesBody.textContent || "";
    }
    const fullText = notesBody?.dataset.fullText || "";
    if (notesBody) notesBody.textContent = fullText;

    if (tryAppendCard(card)) return;

    if (container.children.length) {
      addPage();
    }

    splitCardIntoPages(card, fullText);
  });

  pages.forEach((page, index) => {
    if (index < pages.length - 1) {
      page.classList.add("has-continue");
    }
  });
}

    function paginateTagExplanationCards(root = document) {
  const scope = root.querySelector(".doc-shell") || root;
  const source = scope.querySelector("[data-tag-explanation-source]");
  if (!source) return;

  scope.querySelectorAll(".tag-explanation-page").forEach((page) => page.remove());

  const parent = source.parentNode;
  const anchor = source.nextSibling;
  const itemTemplates = [...source.querySelectorAll(".tag-explanation-item")];
  if (!itemTemplates.length) return;

  const createPage = () => {
    const page = document.createElement("section");
    page.className = "page tag-explanation-page";
    page.innerHTML =
      '<div class="tag-explanation-cards"></div><div class="page-continue-note">Section continues on next page.</div>';
    return page;
  };

  const pages = [];
  let currentPage = createPage();
  parent.insertBefore(currentPage, anchor);
  pages.push(currentPage);
  let container = currentPage.querySelector(".tag-explanation-cards");

  const addPage = () => {
    currentPage = createPage();
    parent.insertBefore(currentPage, anchor);
    pages.push(currentPage);
    container = currentPage.querySelector(".tag-explanation-cards");
  };

  const tryAppendCard = (card) => {
    container.appendChild(card);
    const fits = container.scrollHeight <= container.clientHeight + 0.5;
    if (!fits) container.removeChild(card);
    return fits;
  };

  const markContinuation = (card) => {
    card.classList.add("continued");
    const title = card.querySelector(".tag-explanation-title");
    if (title && !title.querySelector(".tag-explanation-cont")) {
      const marker = document.createElement("span");
      marker.className = "tag-explanation-cont";
      marker.textContent = "(cont.)";
      title.appendChild(marker);
    }
  };

  const splitCardIntoPages = (baseCard, fullText) => {
    let remaining = String(fullText || "").trim();
    if (!remaining) {
      container.appendChild(baseCard);
      return;
    }

    let first = true;
    while (remaining) {
      const card = first ? baseCard : baseCard.cloneNode(true);
      if (!first) {
        markContinuation(card);
      }

      const body = card.querySelector(".tag-explanation-body");
      if (!body) {
        container.appendChild(card);
        return;
      }

      const tokens = remaining.match(/\\S+|\\s+/g) || [];
      container.appendChild(card);

      let low = 1;
      let high = tokens.length;
      let best = 0;

      while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        body.textContent = tokens.slice(0, mid).join("");
        if (container.scrollHeight <= container.clientHeight + 0.5) {
          best = mid;
          low = mid + 1;
        } else {
          high = mid - 1;
        }
      }

      if (best <= 0) {
        body.textContent = remaining;
        return;
      }

      body.textContent = tokens.slice(0, best).join("");
      remaining = tokens.slice(best).join("").trimStart();

      if (remaining) {
        addPage();
        first = false;
      } else {
        return;
      }
    }
  };

  itemTemplates.forEach((template) => {
    const card = template.cloneNode(true);
    const body = card.querySelector(".tag-explanation-body");
    if (body && !body.dataset.fullText) {
      body.dataset.fullText = body.textContent || "";
    }
    const fullText = body?.dataset.fullText || "";
    if (body) body.textContent = fullText;

    if (tryAppendCard(card)) return;

    if (container.children.length) {
      addPage();
      if (tryAppendCard(card)) return;
    }

    splitCardIntoPages(card, fullText);
  });

  pages.forEach((page, index) => {
    if (index < pages.length - 1) {
      page.classList.add("has-continue");
    }
  });
}

function fitSectionContent(section) {
      const container = section.querySelector(".section-body");
      const content = container?.querySelector(".fit-content");
      if (!container || !content) return;

      content.style.transform = "scale(1)";
      content.style.width = "100%";

      const availableHeight = container.clientHeight;
      const neededHeight = content.scrollHeight;
      if (!availableHeight || !neededHeight) return;

      function applyScale(scale) {
        if (scale < 1) {
          content.style.transform = "scale(" + scale + ")";
          content.style.width = (100 / scale) + "%";
        } else {
          content.style.transform = "scale(1)";
          content.style.width = "100%";
        }
      }

      let scale = 1;
      const heightScale = availableHeight / neededHeight;

      // Fill vertical space first.
      if (heightScale < 1) {
        scale = heightScale;
      }

      applyScale(scale);

    }

    function fitAllSummarySections() {
      fitTagFonts(document);
      document.querySelectorAll(".summary-page .section-block").forEach(fitSectionContent);
    }

    window.addEventListener("load", () => {
      fitAllSummarySections();
      paginateReaderCards();
      paginateTagExplanationCards();
      const finish = () => {
        fitAllSummarySections();
        paginateReaderCards();
        paginateTagExplanationCards();
        window.focus();
        window.print();
      };
      if (document.fonts?.ready) {
        document.fonts.ready.then(() => setTimeout(finish, 80));
      } else {
        setTimeout(finish, 80);
      }
    });
  })();
`;

const csvFileInput = document.getElementById("csvFile");
const generateBtn = document.getElementById("generateBtn");
const studentSelect = document.getElementById("studentSelect");
const lsatInput = document.getElementById("lsatInput");
const gpaInput = document.getElementById("gpaInput");
const kjdSelect = document.getElementById("kjdSelect");
const urmSelect = document.getElementById("urmSelect");
const summaryInput = document.getElementById("summaryInput");
const nextStepsInput = document.getElementById("nextStepsInput");
const inputHintEl = document.getElementById("inputHint");
const downloadPdfBtn = document.getElementById("downloadPdfBtn");
const statusEl = document.getElementById("status");
const validationEl = document.getElementById("validation");
const previewRoot = document.getElementById("previewRoot");

csvFileInput.addEventListener("change", onCsvSelected);
generateBtn.addEventListener("click", onGenerateDocuments);
studentSelect.addEventListener("change", onStudentSelectionChange);
lsatInput.addEventListener("input", onManualInputChange);
gpaInput.addEventListener("input", onManualInputChange);
kjdSelect.addEventListener("change", onManualInputChange);
urmSelect.addEventListener("change", onManualInputChange);
summaryInput.addEventListener("input", onManualInputChange);
nextStepsInput.addEventListener("input", onManualInputChange);
downloadPdfBtn.addEventListener("click", onDownloadCurrentStudentPdf);
window.addEventListener("resize", onWindowResize);

function setStatus(message) {
  statusEl.textContent = message;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeProfileKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

const READER_PROFILE_NAME_MAP = new Map(
  READER_PROFILES.flatMap((profile) => {
    const keys = [
      profile.fullName,
      ...(profile.aliases || []),
    ]
      .map(normalizeProfileKey)
      .filter(Boolean);
    return keys.map((key) => [key, profile]);
  })
);

const READER_PROFILE_FIRSTNAME_MAP = READER_PROFILES.reduce((acc, profile) => {
  const key = normalizeProfileKey(profile.firstName);
  if (!key) return acc;
  if (!acc.has(key)) acc.set(key, []);
  acc.get(key).push(profile);
  return acc;
}, new Map());

function getReaderProfile(name) {
  const normalized = normalizeProfileKey(name);
  if (!normalized) return null;
  const direct = READER_PROFILE_NAME_MAP.get(normalized);
  if (direct) return direct;

  if (!normalized.includes(" ")) {
    const candidates = READER_PROFILE_FIRSTNAME_MAP.get(normalized) || [];
    if (candidates.length === 1) return candidates[0];
  }
  return null;
}

function getReaderInitials(label) {
  const trimmed = String(label || "").trim();
  if (!trimmed) return "R";
  return trimmed
    .split(/\s+/)
    .map((part) => part[0])
    .slice(0, 2)
    .join("")
    .toUpperCase();
}

function getReaderDisplayName(rawReviewerName, fallbackLabel = "") {
  const trimmed = String(rawReviewerName || "").trim();
  const profile = getReaderProfile(trimmed);
  return String(profile?.fullName || trimmed || fallbackLabel || "").trim();
}

function getReaderSlotClass(slotLabel) {
  const normalized = String(slotLabel || "").trim();
  return /^(1|2|3)$/.test(normalized) ? `reader-slot-${normalized}` : "";
}

function showValidationMessages() {
  const blocks = [];
  if (state.errors.length) {
    blocks.push(
      `<div class="error"><strong>Errors:</strong> ${escapeHtml(
        state.errors.join(" | ")
      )}</div>`
    );
  }
  if (state.warnings.length) {
    blocks.push(
      `<div class="warn"><strong>Warnings:</strong> ${escapeHtml(
        state.warnings.join(" | ")
      )}</div>`
    );
  }
  validationEl.innerHTML = blocks.length ? blocks.join("") : "No validation issues.";
}

function clearGeneratedResults({ resetSelector } = { resetSelector: false }) {
  state.reports = [];
  state.currentReport = null;
  state.warnings = [];
  state.errors = [];
  previewRoot.innerHTML =
    '<p class="placeholder">Upload a CSV and generate documents to preview output.</p>';
  if (resetSelector) {
    studentSelect.disabled = true;
    studentSelect.innerHTML = "<option>Upload CSV first</option>";
  }
  downloadPdfBtn.disabled = true;
}

function parseCsv(csvText) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < csvText.length; i += 1) {
    const ch = csvText[i];
    const next = csvText[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        cell += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && ch === ",") {
      row.push(cell);
      cell = "";
      continue;
    }

    if (!inQuotes && (ch === "\n" || ch === "\r")) {
      if (ch === "\r" && next === "\n") i += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }

    cell += ch;
  }

  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }

  return rows.filter((r) => r.some((c) => c.trim().length > 0));
}

function readCsvRows(csvText) {
  const allRows = parseCsv(csvText);
  if (!allRows.length) throw new Error("CSV is empty.");

  const normalizeHeaderKey = (value) =>
    String(value || "")
      .replace(/^\uFEFF/, "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
  const normalizedHeaderMap = new Map(
    REQUIRED_HEADERS.map((header) => [
      normalizeHeaderKey(header),
      header,
    ])
  );
  const headers = allRows[0].map((h) => {
    const normalized = normalizeHeaderKey(h);
    return normalizedHeaderMap.get(normalized) || h.replace(/^\uFEFF/, "").trim();
  });
  if (!headers.length || headers.every((h) => h.length === 0)) {
    throw new Error("CSV is missing a header row.");
  }

  const rows = allRows.slice(1).map((rowValues) => {
    const row = {};
    headers.forEach((header, index) => {
      row[header] = (rowValues[index] ?? "").trim();
    });
    return row;
  });

  return { headers, rows };
}

function validateHeaders(headers) {
  return REQUIRED_HEADERS.filter((header) => !headers.includes(header));
}

function groupRowsByFile(rows) {
  const groups = new Map();
  let blankFileRows = 0;

  rows.forEach((row) => {
    const file = (row.File || "").trim();
    if (!file) {
      blankFileRows += 1;
      return;
    }
    if (!groups.has(file)) groups.set(file, []);
    groups.get(file).push(row);
  });

  return { groups, blankFileRows };
}

function shuffleReaders(rows) {
  const shuffled = [...rows];
  for (let i = shuffled.length - 1; i > 0; i -= 1) {
    const j = Math.floor(Math.random() * (i + 1));
    [shuffled[i], shuffled[j]] = [shuffled[j], shuffled[i]];
  }
  return shuffled;
}

function normalizeRubricValue(value, fileName, fieldLabel) {
  const trimmed = String(value || "").trim();
  if (!trimmed) return null;
  if (!(trimmed in RUBRIC_DELTA)) {
    state.warnings.push(`${fileName}: unknown ${fieldLabel} value "${trimmed}"`);
    return null;
  }
  return trimmed;
}

function computeStarsFromRatings(ratings) {
  const filtered = ratings.filter(Boolean);
  if (!filtered.length) return null;
  let stars = 4;
  filtered.forEach((rating) => {
    stars += RUBRIC_DELTA[rating] ?? 0;
  });
  if (stars > 5) stars = 5;
  if (stars < 1) stars = 1;
  return stars;
}

function parseTags(csvTagString) {
  return String(csvTagString || "")
    .split(",")
    .map((tag) => tag.trim())
    .filter(Boolean);
}

function canonicalizeTagName(tag) {
  const key = normalizeTagKey(tag);
  return TAG_CANONICAL_NAME_MAP.get(key) || "";
}

function mapBandValue(value, fileName, fieldLabel) {
  const trimmed = String(value || "").trim();
  if (!trimmed) return "";
  const normalized = trimmed.toUpperCase().replace(/\s+/g, "");
  const aliasMap = {
    "3": "T3",
    "6": "T6",
    "14": "T14",
    "20": "T20",
    "30": "T30",
    "50": "T50",
    "75": "T75",
    "100": "T100",
    "100+": "T100+",
    T3: "T3",
    T6: "T6",
    T14: "T14",
    T20: "T20",
    T25: "T20",
    T30: "T30",
    T50: "T50",
    T75: "T75",
    T100: "T100",
    "T100+": "T100+",
  };
  const canonical = aliasMap[normalized];
  if (canonical && BAND_LABELS[canonical]) return canonical;
  state.warnings.push(`${fileName}: unknown ${fieldLabel} value "${trimmed}"`);
  return `Unknown band: ${trimmed}`;
}

function mapSoftsValue(value, fileName) {
  const trimmed = String(value || "").trim().toUpperCase();
  if (!trimmed) return null;
  const match = /^T([1-4])$/.exec(trimmed);
  if (!match) {
    state.warnings.push(`${fileName}: unknown Softs value "${trimmed}"`);
    return null;
  }
  return Number(match[1]);
}

function computeSoftsDisplay(rows, fileName) {
  const values = rows
    .map((row) => mapSoftsValue(row.Softs, fileName))
    .filter((value) => value !== null);
  if (!values.length) return null;

  const counts = values.reduce((acc, value) => {
    acc.set(value, (acc.get(value) || 0) + 1);
    return acc;
  }, new Map());

  if (counts.size === 1) {
    const [onlyValue] = counts.keys();
    return `T${onlyValue}`;
  }

  if (counts.size === 2) {
    const sortedCounts = [...counts.entries()].sort((a, b) => b[1] - a[1]);
    return `T${sortedCounts[0][0]}/T${sortedCounts[1][0]}`;
  }

  const average =
    values.reduce((sum, value) => sum + value, 0) / values.length;
  return `T${average.toFixed(1)}`;
}

function uniqueNonEmpty(values) {
  return [...new Set(values.map((v) => String(v || "").trim()).filter(Boolean))];
}

function buildStudentReport(fileName, rows, manual) {
  const unknownTags = new Set();
  const summaryRatings = {
    whyLaw: [],
    thrive: [],
    contribute: [],
    know: [],
  };
  const randomizedRows = shuffleReaders(rows);
  const labeledReaders = randomizedRows.map((row, index) => {
    const rawTags = parseTags(row.Tags);
    const tags = [];
    rawTags.forEach((rawTag) => {
      const canonicalTag = canonicalizeTagName(rawTag);
      if (canonicalTag && TAG_NAME_SET.has(canonicalTag)) {
        if (!tags.includes(canonicalTag)) tags.push(canonicalTag);
      } else if (rawTag) {
        unknownTags.add(rawTag);
      }
    });
    return {
      row,
      label: getReaderDisplayName(row.Reviewer, `Reader ${index + 1}`),
      badgeLabel: String(index + 1),
      tags,
    };
  });

  const normalizeBandDisplay = (value) => {
    const trimmed = String(value || "").trim();
    if (!trimmed || trimmed.startsWith("Unknown band:")) return "—";
    return trimmed;
  };

  const readers = labeledReaders.map(({ row, label, badgeLabel, tags }) => {
    const ratingLabels = {
      whyLaw: normalizeRubricValue(row["Why Law?"], fileName, "Why Law?"),
      thrive: normalizeRubricValue(row["Thrive?"], fileName, "Thrive?"),
      contribute: normalizeRubricValue(row["Contribute?"], fileName, "Contribute?"),
      know: normalizeRubricValue(row["Know?"], fileName, "Know?"),
    };
    if (ratingLabels.whyLaw) summaryRatings.whyLaw.push(ratingLabels.whyLaw);
    if (ratingLabels.thrive) summaryRatings.thrive.push(ratingLabels.thrive);
    if (ratingLabels.contribute) summaryRatings.contribute.push(ratingLabels.contribute);
    if (ratingLabels.know) summaryRatings.know.push(ratingLabels.know);

    return {
      label,
      badgeLabel,
      notes: row.Notes || "",
      anythingElse: row["Anything Else?"] || "",
      rawReviewer: row.Reviewer || "",
      ratingLabels,
      ratingStars: {
        whyLaw: computeStarsFromRatings([ratingLabels.whyLaw]),
        thrive: computeStarsFromRatings([ratingLabels.thrive]),
        contribute: computeStarsFromRatings([ratingLabels.contribute]),
        know: computeStarsFromRatings([ratingLabels.know]),
      },
      bands: {
        reach: normalizeBandDisplay(mapBandValue(row.Reach, fileName, "Reach")),
        target: normalizeBandDisplay(mapBandValue(row.Target, fileName, "Target")),
        safety: normalizeBandDisplay(mapBandValue(row.Safety, fileName, "Safety")),
      },
      softs: (() => {
        const softValue = mapSoftsValue(row.Softs, fileName);
        return softValue ? `T${softValue}` : "—";
      })(),
      tags,
    };
  });

  const activeTags = new Set();
  const tagReaderMap = new Map();
  labeledReaders.forEach(({ badgeLabel, tags }) => {
    tags.forEach((tag) => {
      if (TAG_NAME_SET.has(tag)) {
        activeTags.add(tag);
        if (!tagReaderMap.has(tag)) tagReaderMap.set(tag, new Set());
        tagReaderMap.get(tag).add(badgeLabel);
      } else {
        unknownTags.add(tag);
      }
    });
  });

  if (unknownTags.size) {
    state.warnings.push(
      `${fileName}: unknown tag(s) ${[...unknownTags].sort().join(", ")}`
    );
  }

  return {
    fileName,
    manual,
    readers,
    summaryStars: {
      whyLaw: computeStarsFromRatings(summaryRatings.whyLaw),
      thrive: computeStarsFromRatings(summaryRatings.thrive),
      contribute: computeStarsFromRatings(summaryRatings.contribute),
      know: computeStarsFromRatings(summaryRatings.know),
    },
    bands: {
      reach: uniqueNonEmpty(rows.map((row) => mapBandValue(row.Reach, fileName, "Reach"))),
      target: uniqueNonEmpty(rows.map((row) => mapBandValue(row.Target, fileName, "Target"))),
      safety: uniqueNonEmpty(rows.map((row) => mapBandValue(row.Safety, fileName, "Safety"))),
    },
    softsDisplay: computeSoftsDisplay(rows, fileName),
    activeTags,
    tagReaderMap,
  };
}

function renderStarSvg(type, idSuffix) {
  if (type === "half") {
    return `<svg class="star half" viewBox="0 0 24 24" aria-hidden="true">
      <defs>
        <linearGradient id="half-fill-${idSuffix}" x1="0" x2="1" y1="0" y2="0">
          <stop offset="50%" stop-color="#f4b400"></stop>
          <stop offset="50%" stop-color="#eef2f7"></stop>
        </linearGradient>
      </defs>
      <polygon fill="url(#half-fill-${idSuffix})" points="12,2 15.3,8.8 22.8,9.8 17.3,15.1 18.7,22.5 12,18.8 5.3,22.5 6.7,15.1 1.2,9.8 8.7,8.8"></polygon>
    </svg>`;
  }

  return `<svg class="star ${type}" viewBox="0 0 24 24" aria-hidden="true">
    <polygon points="12,2 15.3,8.8 22.8,9.8 17.3,15.1 18.7,22.5 12,18.8 5.3,22.5 6.7,15.1 1.2,9.8 8.7,8.8"></polygon>
  </svg>`;
}

function renderStarRow(label, starCount) {
  if (starCount === null) {
    return `<div class="rating-row"><div class="label">${escapeHtml(label)}</div><div class="small">No rating submitted</div></div>`;
  }

  const count = Math.max(1, Math.min(5, Math.round(starCount)));
  const stars = Array.from({ length: count }, (_, idx) =>
    renderStarSvg("filled", `${label}-full-${idx}`)
  ).join("");
  return `<div class="rating-row">
    <div class="label">${escapeHtml(label)}</div>
    <div class="stars-wrap">
      <div class="stars">${stars}</div>
    </div>
  </div>`;
}

function renderBandList(label, values) {
  if (!values.length) {
    return `<div class="row"><div class="label">${escapeHtml(label)}</div><div class="value">No submissions</div></div>`;
  }

  return `<div class="row">
    <div class="label">${escapeHtml(label)}</div>
    <ul class="band-list">${values
      .map((v) => `<li>${escapeHtml(BAND_LABELS[v] || v)}</li>`)
      .join("")}</ul>
  </div>`;
}

function renderTagBadges(readerLabels) {
  if (!readerLabels.length) return "";
  return `<span class="tag-badges">${readerLabels
    .map((label) => {
      const slotClass = getReaderSlotClass(label);
      return `<span class="tag-badge${slotClass ? ` ${slotClass}` : ""}">${escapeHtml(label)}</span>`;
    })
    .join("")}</span>`;
}

function renderTagGrid(activeTags, tagReaderMap) {
  return `<div class="tag-grid">
    ${DISPLAY_TAG_DEFINITIONS.map((tag) => {
      const isActive = activeTags.has(tag.name);
      const readerLabels = isActive
        ? [...(tagReaderMap.get(tag.name) || new Set())].sort((a, b) =>
            a.localeCompare(b)
          )
        : [];
      const className = isActive
        ? tag.polarity === "positive"
          ? "tag-pill active-positive"
          : "tag-pill active-negative"
        : "tag-pill inactive";
      return `<div class="${className}"><span class="tag-text">${escapeHtml(
        tag.name
      )}</span>${renderTagBadges(readerLabels)}</div>`;
    }).join("")}
  </div>`;
}

function renderCompactStars(starCount) {
  if (starCount === null) {
    return `<span class="compact-stars"></span>`;
  }

  const count = Math.max(1, Math.min(5, Math.round(starCount)));
  const stars = Array.from({ length: count }, (_, idx) =>
    renderStarSvg("filled", `compact-full-${idx}`)
  ).join("");
  return `<span class="compact-stars">${stars}</span>`;
}

function renderBandRow(label, className, values) {
  const rankedValues = values
    .map((value) => BAND_ORDER.indexOf(value))
    .filter((index) => index >= 0)
    .sort((a, b) => a - b);
  const minIndex = rankedValues.length ? rankedValues[0] : null;
  const maxIndex = rankedValues.length ? rankedValues[rankedValues.length - 1] : null;

  return `<div class="band-row row-${className}">
    <div class="band-name ${className}" style="grid-column: 1; grid-row: 1;">${escapeHtml(label)}</div>
    ${
      minIndex !== null && maxIndex !== null
        ? `<div class="band-range range-${className}" style="grid-column: ${minIndex + 2} / ${maxIndex + 3}; grid-row: 1;"></div>`
        : ""
    }
    ${BAND_ORDER.map((value, idx) => {
      const inRange =
        minIndex !== null && maxIndex !== null && idx >= minIndex && idx <= maxIndex;
      return `<div class="band-chip${inRange ? " in-range" : ""}" style="grid-column: ${idx + 2}; grid-row: 1;">${escapeHtml(value)}</div>`;
    }).join("")}
  </div>`;
}

function renderTagGridFourColumns(activeTags, tagReaderMap) {
  return `<div class="design-tag-grid">
    ${DISPLAY_TAG_DEFINITIONS.map((tag) => {
      const isActive = activeTags.has(tag.name);
      const readerLabels = isActive
        ? [...(tagReaderMap.get(tag.name) || new Set())].sort((a, b) =>
            a.localeCompare(b)
          )
        : [];
      const className = isActive
        ? tag.polarity === "positive"
          ? "tag-pill active-positive"
          : "tag-pill active-negative"
        : "tag-pill inactive";
      return `<div class="${className}"><span class="tag-text">${escapeHtml(
        tag.name
      )}</span>${renderTagBadges(readerLabels)}</div>`;
    }).join("")}
  </div>`;
}

function renderReaders(readers) {
  return readers
    .map(
      (reader) => `<article class="reader-card">
        <h3 class="reader-title">${escapeHtml(reader.label)}</h3>
        <div class="label">Notes</div>
        <div class="notes-box">${escapeHtml(reader.notes)}</div>
        <div class="label">Anything Else?</div>
        <div class="notes-box">${escapeHtml(reader.anythingElse)}</div>
      </article>`
    )
    .join("");
}

function splitBioIntoPoints(bioText) {
  const normalized = String(bioText || "")
    .replace(/\s+/g, " ")
    .trim();
  if (!normalized) return [];
  return (normalized.match(/[^.!?]+[.!?]?/g) || [])
    .map((part) => part.trim())
    .filter(Boolean);
}

function renderReaderSummaryBio(bioText) {
  const points = splitBioIntoPoints(bioText);
  if (!points.length) {
    return '<p class="reader-bio">—</p>';
  }
  return `<ul class="reader-bio-list">${points
    .map((point) => `<li>${escapeHtml(point)}</li>`)
    .join("")}</ul>`;
}

function renderReaderRatingRow(label, ratingValue) {
  const display = ratingValue ? String(ratingValue) : "—";
  const toneClassMap = {
    "Strongly Agree": "is-positive-strong",
    Agree: "is-positive",
    Disagree: "is-negative",
    "Strongly Disagree": "is-negative-strong",
  };
  const toneClass = ratingValue ? toneClassMap[String(ratingValue)] || "" : "";
  const pillClass = ratingValue
    ? `reader-rating-pill${toneClass ? ` ${toneClass}` : ""}`
    : "reader-rating-pill empty";
  return `<div class="reader-rating-row">
    <div class="reader-rating-label">${escapeHtml(label)}</div>
    <span class="${pillClass}">${escapeHtml(display)}</span>
  </div>`;
}

function renderReaderBandRow(label, value) {
  const display = value && value !== "—" ? value : "—";
  return `<div class="reader-band-row">
    <span class="reader-band-label">${escapeHtml(label)}</span>
    <span class="reader-band-value">${escapeHtml(display)}</span>
  </div>`;
}

function renderReaderTags(tags) {
  const visibleTags = tags.filter((tag) => !HIDDEN_TAG_NAMES.has(tag));
  if (!visibleTags.length) {
    return `<div class="reader-tag-empty">—</div>`;
  }
  return visibleTags
    .map((tag) => {
      const polarity = TAG_POLARITY_MAP.get(tag);
      const className =
        polarity === "positive"
          ? "reader-tag-pill active-positive"
          : polarity === "negative"
            ? "reader-tag-pill active-negative"
            : "reader-tag-pill inactive";
      return `<div class="${className}"><span class="reader-tag-text">${escapeHtml(
        tag
      )}</span></div>`;
    })
    .join("");
}

function getReaderNotesText(reader) {
  const notes = String(reader.notes || "").trim();
  const anythingElse = String(reader.anythingElse || "").trim();
  return notes && anythingElse
    ? `${notes}\n\n${anythingElse}`
    : notes || anythingElse || "—";
}

function renderReaderNotesBlock(text) {
  return `<div class="reader-notes">
    <div class="label">Notes</div>
    <div class="reader-notes-body">${escapeHtml(text)}</div>
  </div>`;
}

function renderReaderCard(reader) {
  const notesText = getReaderNotesText(reader);
  const slotClass = getReaderSlotClass(reader.badgeLabel);
  return `
    <article class="reader-card${slotClass ? ` ${slotClass}` : ""}">
      <div class="reader-section-title">${escapeHtml(reader.label)}</div>
      <div class="reader-top-grid">
        <div class="reader-col ratings">
          ${renderReaderRatingRow(
            "I understand the candidate's reason for seeking a law degree.",
            reader.ratingLabels.whyLaw
          )}
          ${renderReaderRatingRow(
            "This candidate will thrive in law school.",
            reader.ratingLabels.thrive
          )}
          ${renderReaderRatingRow(
            "This candidate would positively contribute to a law school community.",
            reader.ratingLabels.contribute
          )}
          ${renderReaderRatingRow(
            "i feel like I know this candidate.",
            reader.ratingLabels.know
          )}
        </div>
        <div class="reader-col bands">
          ${renderReaderBandRow("Reach", reader.bands.reach)}
          ${renderReaderBandRow("Target", reader.bands.target)}
          ${renderReaderBandRow("Safety", reader.bands.safety)}
          ${renderReaderBandRow("Softs", reader.softs)}
        </div>
        <div class="reader-col tags">
          ${renderReaderTags(reader.tags)}
        </div>
      </div>
      ${renderReaderNotesBlock(notesText)}
    </article>
  `;
}

function renderReaderPages(report) {
  return `
    <div class="reader-cards-source" data-reader-source>
      ${report.readers.map((reader) => renderReaderCard(reader)).join("")}
    </div>
  `;
}

function renderSummaryNextStepsPage(report) {
  const summaryText = String(report.manual.summary || "").trim() || "—";
  const nextStepsText = String(report.manual.nextSteps || "").trim() || "—";
  return `
    <section class="page reader-summary-page">
      <div class="reader-summary-grid">
        <div class="reader-footer-box">
          <div class="reader-footer-title">Summary</div>
          <div class="reader-footer-body">${escapeHtml(summaryText)}</div>
        </div>
        <div class="reader-footer-box">
          <div class="reader-footer-title">Recommended Next Steps</div>
          <div class="reader-footer-body">${escapeHtml(nextStepsText)}</div>
        </div>
      </div>
    </section>
  `;
}

function renderWaitingPage() {
  return `
    <section class="page waiting-page">
      <p class="waiting-copy">waiting on design</p>
    </section>
  `;
}

function renderTagExplanationItem(tag) {
  const polarity = TAG_POLARITY_MAP.get(tag);
  const className =
    polarity === "positive"
      ? "tag-pill active-positive"
      : polarity === "negative"
        ? "tag-pill active-negative"
        : "tag-pill inactive";
  return `<article class="tag-explanation-item">
    <div class="tag-explanation-title">
      <span class="${className} tag-explanation-pill">
        <span class="tag-text">${escapeHtml(tag)}</span>
      </span>
    </div>
    <div class="tag-explanation-body">${escapeHtml(
      TAG_DESCRIPTION_MAP.get(tag) || "Description coming soon."
    )}</div>
  </article>`;
}

function renderTagExplanationPages(report) {
  const tags = [...(report.activeTags || new Set())]
    .filter((tag) => !HIDDEN_TAG_NAMES.has(tag))
    .sort((a, b) => a.localeCompare(b));
  if (!tags.length) {
    return `
      <section class="page tag-explanation-page">
        <div class="tag-explanation-empty">No tags selected.</div>
      </section>
    `;
  }
  return `
    <div class="tag-explanation-source" data-tag-explanation-source>
      ${tags.map((tag) => renderTagExplanationItem(tag)).join("")}
    </div>
  `;
}

function renderStudentDocument(report) {
  const summaryReaders = report.readers
    .map((reader) => ({ reader, slot: Number.parseInt(reader.badgeLabel, 10) || 999 }))
    .sort((a, b) => a.slot - b.slot)
    .map(({ reader }) => reader);

  return `
    <section class="page summary-page">
      <div class="summary-banner">Summary</div>
      <div class="section-block basics-section">
        <div class="section-body">
          <div class="fit-content fit-basics">
            <div class="section-title">Basics</div>
            <div class="basics-grid">
              <div class="basics-item">
                <span class="basics-label">LSAT:</span>
                <span class="basics-value">${escapeHtml(report.manual.lsat || "—")}</span>
              </div>
              <span class="basics-divider">|</span>
              <div class="basics-item">
                <span class="basics-label">GPA:</span>
                <span class="basics-value">${escapeHtml(report.manual.gpa || "—")}</span>
              </div>
              <span class="basics-divider">|</span>
              <div class="basics-item">
                <span class="basics-label">Other:</span>
                <span class="basics-value">${escapeHtml(report.manual.otherText || "—")}</span>
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="section-block readers-section">
        <div class="section-body">
          <div class="fit-content fit-readers">
          <div class="section-title">Readers</div>
          <div class="readers-card">
            <div class="reader-grid">
              ${summaryReaders
                .map((reader) => {
                  const profile = getReaderProfile(reader.rawReviewer);
                  const name = profile?.fullName || reader.rawReviewer || reader.label;
                  const slotClass = getReaderSlotClass(reader.badgeLabel);
                  const avatarContent = profile?.headshotUrl
                    ? `<img src="${escapeHtml(profile.headshotUrl)}" alt="${escapeHtml(name)}" />`
                    : escapeHtml(getReaderInitials(name));
                  const bio =
                    profile?.bio ||
                    "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.";
                  return `<div class="reader-col">
                    <div class="avatar${profile?.headshotUrl ? " has-photo" : ""}${slotClass ? ` ${slotClass}` : ""}">${avatarContent}</div>
                    <div>
                      <h3 class="reader-name">${escapeHtml(name)}</h3>
                      ${renderReaderSummaryBio(bio)}
                    </div>
                  </div>`;
                })
                .join("")}
            </div>
          </div>
          </div>
        </div>
      </div>
      <div class="section-block key-takeaways-section">
        <div class="section-body">
          <div class="fit-content fit-takeaways">
            <div class="section-title">Key Takeaways</div>
            <div class="takeaways-card">
              <div class="takeaways-row takeaways-row-top">
                <div class="takeaway-item">
                  <div class="takeaway-title">Why Law</div>
                  ${renderCompactStars(report.summaryStars.whyLaw)}
                </div>
                <div class="takeaway-item">
                  <div class="takeaway-title">Readiness</div>
                  ${renderCompactStars(report.summaryStars.thrive)}
                </div>
                <div class="takeaway-item takeaway-softs">
                  <div class="takeaway-title">Softs</div>
                  <div class="takeaway-softs-value">${escapeHtml(report.softsDisplay || "—")}</div>
                </div>
              </div>
              <div class="takeaways-row takeaways-row-bottom">
                <div class="takeaway-item">
                  <div class="takeaway-title">Perspective</div>
                  ${renderCompactStars(report.summaryStars.contribute)}
                </div>
                <div class="takeaway-item">
                  <div class="takeaway-title">Personality</div>
                  ${renderCompactStars(report.summaryStars.know)}
                </div>
              </div>
              <div class="takeaways-bands">
                ${renderBandRow("Reach", "reach", report.bands.reach)}
                ${renderBandRow("Target", "target", report.bands.target)}
                ${renderBandRow("Safety", "safety", report.bands.safety)}
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="section-block tags-section">
        <div class="section-body">
          <div class="fit-content fit-tags">
          <div class="section-title">Tags</div>
          <div class="tags-card">
            <div class="tags-grid-wrap">
              ${renderTagGridFourColumns(report.activeTags, report.tagReaderMap)}
            </div>
          </div>
          </div>
        </div>
      </div>
    </section>

    ${renderReaderPages(report)}
    ${renderSummaryNextStepsPage(report)}
    ${renderTagExplanationPages(report)}
  `;
}

function getDefaultStudentInput() {
  return {
    lsat: "",
    gpa: "",
    kjd: "Not KJD",
    urm: "Non-URM",
    summary: "",
    nextSteps: "",
  };
}

function setManualControlsEnabled(enabled) {
  lsatInput.disabled = !enabled;
  gpaInput.disabled = !enabled;
  kjdSelect.disabled = !enabled;
  urmSelect.disabled = !enabled;
  summaryInput.disabled = !enabled;
  nextStepsInput.disabled = !enabled;
}

function getSelectedFile() {
  return String(studentSelect.value || "").trim();
}

function ensureStudentInput(fileName) {
  if (!state.studentInputByFile[fileName]) {
    state.studentInputByFile[fileName] = getDefaultStudentInput();
  }
  return state.studentInputByFile[fileName];
}

function syncManualFormFromSelection() {
  const fileName = getSelectedFile();
  const hasSelection = Boolean(fileName);
  setManualControlsEnabled(hasSelection);
  if (!hasSelection) {
    lsatInput.value = "";
    gpaInput.value = "";
    kjdSelect.value = "Not KJD";
    urmSelect.value = "Non-URM";
    summaryInput.value = "";
    nextStepsInput.value = "";
    inputHintEl.textContent = "";
    return;
  }

  const manual = ensureStudentInput(fileName);
  lsatInput.value = manual.lsat;
  gpaInput.value = manual.gpa;
  kjdSelect.value = manual.kjd;
  urmSelect.value = manual.urm;
  summaryInput.value = manual.summary || "";
  nextStepsInput.value = manual.nextSteps || "";
  validateManualInputs();
}

function setStudentSelectorOptions() {
  studentSelect.innerHTML = "";
  if (!state.availableStudents.length) {
    studentSelect.innerHTML = "<option>No valid student names found</option>";
    studentSelect.disabled = true;
    setManualControlsEnabled(false);
    return;
  }

  state.availableStudents.forEach((fileName) => {
    const option = document.createElement("option");
    option.value = fileName;
    option.textContent = fileName;
    studentSelect.appendChild(option);
  });
  studentSelect.disabled = false;
  syncManualFormFromSelection();
}

function renderPreviewFromCurrent() {
  const report = state.currentReport;
  if (!report) {
    previewRoot.innerHTML =
      '<p class="placeholder">No document available. Generate documents first.</p>';
    return;
  }
  previewRoot.innerHTML = `<div class="doc-shell">${renderStudentDocument(report)}</div>`;
  fitAllSummarySections(previewRoot);
  paginateReaderCards(previewRoot);
  paginateTagExplanationCards(previewRoot);
  requestAnimationFrame(() => {
    fitAllSummarySections(previewRoot);
    paginateReaderCards(previewRoot);
    paginateTagExplanationCards(previewRoot);
  });
  if (document.fonts?.ready) {
    document.fonts.ready.then(() => {
      fitAllSummarySections(previewRoot);
      paginateReaderCards(previewRoot);
      paginateTagExplanationCards(previewRoot);
    });
  }
}

function validateManualInputs() {
  const warnings = [];
  const lsat = lsatInput.value.trim();
  const gpa = gpaInput.value.trim();

  if (lsat && !/^\d+$/.test(lsat)) {
    warnings.push("LSAT should usually be digits only.");
  }
  if (gpa && !/^\d+(\.\d+)?$/.test(gpa)) {
    warnings.push("GPA should usually look like a decimal number.");
  }

  inputHintEl.textContent = warnings.join(" ");
}

function onManualInputChange() {
  const fileName = getSelectedFile();
  if (!fileName) return;
  const manual = ensureStudentInput(fileName);
  manual.lsat = lsatInput.value.trim();
  manual.gpa = gpaInput.value.trim();
  manual.kjd = kjdSelect.value;
  manual.urm = urmSelect.value;
  manual.summary = summaryInput.value.trim();
  manual.nextSteps = nextStepsInput.value.trim();
  validateManualInputs();
}

function fitTagFonts(root = document) {
  const pills = root.querySelectorAll(".tag-pill");
  pills.forEach((pill) => {
    const textEl = pill.querySelector(".tag-text");
    if (!textEl) return;

    const prevDisplay = textEl.style.display;
    const prevClamp = textEl.style.webkitLineClamp;
    const prevOrient = textEl.style.webkitBoxOrient;

    textEl.style.display = "block";
    textEl.style.webkitLineClamp = "unset";
    textEl.style.webkitBoxOrient = "initial";

    const setFont = (sizePx) => {
      pill.style.setProperty("--tag-font-size", `${sizePx}px`);
    };

    const getLineHeight = () => {
      const computed = window.getComputedStyle(textEl);
      const fontSize = parseFloat(computed.fontSize) || TAG_FONT_BASE_PX;
      let lineHeight = parseFloat(computed.lineHeight);
      if (!lineHeight || Number.isNaN(lineHeight)) {
        lineHeight = fontSize * 1.1;
      }
      return { lineHeight, fontSize };
    };

    setFont(TAG_FONT_BASE_PX);
    let { lineHeight } = getLineHeight();
    let maxHeight = lineHeight * 2 + 0.5;
    let fullHeight = textEl.scrollHeight;

    if (fullHeight > maxHeight) {
      let low = TAG_FONT_MIN_PX;
      let high = TAG_FONT_BASE_PX;
      let best = TAG_FONT_MIN_PX;

      while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        setFont(mid);
        ({ lineHeight } = getLineHeight());
        maxHeight = lineHeight * 2 + 0.5;
        fullHeight = textEl.scrollHeight;

        if (fullHeight <= maxHeight) {
          best = mid;
          low = mid + 1;
        } else {
          high = mid - 1;
        }
      }
      setFont(best);
    }

    textEl.style.display = prevDisplay;
    textEl.style.webkitLineClamp = prevClamp;
    textEl.style.webkitBoxOrient = prevOrient;
  });
}

function paginateReaderCards(root = document) {
  const scope = root.querySelector(".doc-shell") || root;
  const source = scope.querySelector("[data-reader-source]");
  if (!source) return;

  scope.querySelectorAll(".reader-detail-page").forEach((page) => page.remove());

  const parent = source.parentNode;
  const anchor = source.nextSibling;
  const cardTemplates = [...source.querySelectorAll(".reader-card")];
  if (!cardTemplates.length) return;

  const createPage = () => {
    const page = document.createElement("section");
    page.className = "page reader-detail-page";
    page.innerHTML =
      '<div class="reader-cards"></div><div class="page-continue-note">Section continues on next page.</div>';
    return page;
  };

  const pages = [];
  let currentPage = createPage();
  parent.insertBefore(currentPage, anchor);
  pages.push(currentPage);
  let container = currentPage.querySelector(".reader-cards");

  const addPage = () => {
    currentPage = createPage();
    parent.insertBefore(currentPage, anchor);
    pages.push(currentPage);
    container = currentPage.querySelector(".reader-cards");
  };

  const tryAppendCard = (card) => {
    container.appendChild(card);
    const fits = container.scrollHeight <= container.clientHeight + 0.5;
    if (!fits) container.removeChild(card);
    return fits;
  };

  const splitCardIntoPages = (baseCard, fullText) => {
    let remaining = fullText.trim();
    let first = true;
    while (remaining) {
      let card = first ? baseCard : baseCard.cloneNode(true);
      if (!first) {
        card.classList.add("continued");
        const header = card.querySelector(".reader-section-title");
        if (header && !header.textContent.includes("(cont.)")) {
          header.textContent = header.textContent + " (cont.)";
        }
        const topGrid = card.querySelector(".reader-top-grid");
        if (topGrid) topGrid.remove();
      }

      const notesBody = card.querySelector(".reader-notes-body");
      if (!notesBody) {
        container.appendChild(card);
        return;
      }

      const tokens = remaining.match(/\S+|\s+/g) || [];
      container.appendChild(card);

      let low = 1;
      let high = tokens.length;
      let best = 0;

      while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        notesBody.textContent = tokens.slice(0, mid).join("");
        if (container.scrollHeight <= container.clientHeight + 0.5) {
          best = mid;
          low = mid + 1;
        } else {
          high = mid - 1;
        }
      }

      if (best <= 0) {
        notesBody.textContent = remaining;
        return;
      }

      notesBody.textContent = tokens.slice(0, best).join("");
      remaining = tokens.slice(best).join("").trimStart();

      if (remaining) {
        addPage();
        first = false;
      } else {
        return;
      }
    }
  };

  cardTemplates.forEach((template) => {
    const card = template.cloneNode(true);
    const notesBody = card.querySelector(".reader-notes-body");
    if (notesBody && !notesBody.dataset.fullText) {
      notesBody.dataset.fullText = notesBody.textContent || "";
    }
    const fullText = notesBody?.dataset.fullText || "";
    if (notesBody) notesBody.textContent = fullText;

    if (tryAppendCard(card)) return;

    if (container.children.length) {
      addPage();
    }

    splitCardIntoPages(card, fullText);
  });

  pages.forEach((page, index) => {
    if (index < pages.length - 1) {
      page.classList.add("has-continue");
    }
  });
}

function paginateTagExplanationCards(root = document) {
  const scope = root.querySelector(".doc-shell") || root;
  const source = scope.querySelector("[data-tag-explanation-source]");
  if (!source) return;

  scope.querySelectorAll(".tag-explanation-page").forEach((page) => page.remove());

  const parent = source.parentNode;
  const anchor = source.nextSibling;
  const itemTemplates = [...source.querySelectorAll(".tag-explanation-item")];
  if (!itemTemplates.length) return;

  const createPage = () => {
    const page = document.createElement("section");
    page.className = "page tag-explanation-page";
    page.innerHTML =
      '<div class="tag-explanation-cards"></div><div class="page-continue-note">Section continues on next page.</div>';
    return page;
  };

  const pages = [];
  let currentPage = createPage();
  parent.insertBefore(currentPage, anchor);
  pages.push(currentPage);
  let container = currentPage.querySelector(".tag-explanation-cards");

  const addPage = () => {
    currentPage = createPage();
    parent.insertBefore(currentPage, anchor);
    pages.push(currentPage);
    container = currentPage.querySelector(".tag-explanation-cards");
  };

  const tryAppendCard = (card) => {
    container.appendChild(card);
    const fits = container.scrollHeight <= container.clientHeight + 0.5;
    if (!fits) container.removeChild(card);
    return fits;
  };

  const markContinuation = (card) => {
    card.classList.add("continued");
    const title = card.querySelector(".tag-explanation-title");
    if (title && !title.querySelector(".tag-explanation-cont")) {
      const marker = document.createElement("span");
      marker.className = "tag-explanation-cont";
      marker.textContent = "(cont.)";
      title.appendChild(marker);
    }
  };

  const splitCardIntoPages = (baseCard, fullText) => {
    let remaining = String(fullText || "").trim();
    if (!remaining) {
      container.appendChild(baseCard);
      return;
    }

    let first = true;
    while (remaining) {
      const card = first ? baseCard : baseCard.cloneNode(true);
      if (!first) {
        markContinuation(card);
      }

      const body = card.querySelector(".tag-explanation-body");
      if (!body) {
        container.appendChild(card);
        return;
      }

      const tokens = remaining.match(/\S+|\s+/g) || [];
      container.appendChild(card);

      let low = 1;
      let high = tokens.length;
      let best = 0;

      while (low <= high) {
        const mid = Math.floor((low + high) / 2);
        body.textContent = tokens.slice(0, mid).join("");
        if (container.scrollHeight <= container.clientHeight + 0.5) {
          best = mid;
          low = mid + 1;
        } else {
          high = mid - 1;
        }
      }

      if (best <= 0) {
        body.textContent = remaining;
        return;
      }

      body.textContent = tokens.slice(0, best).join("");
      remaining = tokens.slice(best).join("").trimStart();

      if (remaining) {
        addPage();
        first = false;
      } else {
        return;
      }
    }
  };

  itemTemplates.forEach((template) => {
    const card = template.cloneNode(true);
    const body = card.querySelector(".tag-explanation-body");
    if (body && !body.dataset.fullText) {
      body.dataset.fullText = body.textContent || "";
    }
    const fullText = body?.dataset.fullText || "";
    if (body) body.textContent = fullText;

    if (tryAppendCard(card)) return;

    if (container.children.length) {
      addPage();
      if (tryAppendCard(card)) return;
    }

    splitCardIntoPages(card, fullText);
  });

  pages.forEach((page, index) => {
    if (index < pages.length - 1) {
      page.classList.add("has-continue");
    }
  });
}

function fitSectionContent(section) {
  const container = section.querySelector(".section-body");
  const content = container?.querySelector(".fit-content");
  if (!container || !content) return;

  content.style.transform = "scale(1)";
  content.style.width = "100%";

  const availableHeight = container.clientHeight;
  const neededHeight = content.scrollHeight;
  if (!availableHeight || !neededHeight) return;

  function applyScale(scale) {
    if (scale < 1) {
      content.style.transform = `scale(${scale})`;
      content.style.width = `${100 / scale}%`;
    } else {
      content.style.transform = "scale(1)";
      content.style.width = "100%";
    }
  }

  let scale = 1;
  const heightScale = availableHeight / neededHeight;

  // Fill vertical space first.
  if (heightScale < 1) {
    scale = heightScale;
  }

  applyScale(scale);

}

function fitAllSummarySections(root = document) {
  fitTagFonts(root);
  root.querySelectorAll(".summary-page .section-block").forEach((section) => {
    fitSectionContent(section);
  });
}

function onWindowResize() {
  if (fitResizeScheduled) return;
  fitResizeScheduled = true;
  requestAnimationFrame(() => {
    fitResizeScheduled = false;
    fitAllSummarySections(previewRoot);
    paginateReaderCards(previewRoot);
    paginateTagExplanationCards(previewRoot);
  });
}

async function onCsvSelected(event) {
  clearGeneratedResults({ resetSelector: true });
  generateBtn.disabled = true;
  state.parsedRows = [];
  state.groupedRows = new Map();
  state.availableStudents = [];
  state.studentInputByFile = {};
  state.fileName = "";
  setManualControlsEnabled(false);
  inputHintEl.textContent = "";

  const file = event.target.files[0];
  if (!file) {
    setStatus("Waiting for CSV upload.");
    showValidationMessages();
    return;
  }

  try {
    const csvText = await file.text();
    const { headers, rows } = readCsvRows(csvText);
    const missingHeaders = validateHeaders(headers);
    const { groups, blankFileRows } = groupRowsByFile(rows);

    state.fileName = file.name;
    state.parsedRows = rows;
    state.groupedRows = groups;
    state.availableStudents = [...groups.keys()].sort((a, b) => a.localeCompare(b));
    state.errors = missingHeaders.length
      ? [`Missing required header(s): ${missingHeaders.join(", ")}`]
      : [];
    state.warnings = [];
    if (blankFileRows > 0) {
      state.warnings.push(`Skipped ${blankFileRows} row(s) with blank File.`);
    }
    if (!groups.size && rows.length) {
      state.warnings.push(
        "No student names found. Check the File column for student names."
      );
    }

    generateBtn.disabled = state.errors.length > 0 || state.availableStudents.length === 0;
    setStudentSelectorOptions();
    setStatus(
      state.errors.length
        ? `Loaded ${rows.length} row(s), but CSV contract validation failed.`
        : `Loaded ${rows.length} row(s) from ${file.name}. Found ${state.availableStudents.length} student(s).`
    );
  } catch (error) {
    state.errors = [`Could not parse CSV: ${error.message}`];
    generateBtn.disabled = true;
    setStatus("CSV load failed.");
  }

  showValidationMessages();
}

function onGenerateDocuments() {
  clearGeneratedResults();
  if (!state.parsedRows.length) {
    state.errors = ["No CSV rows available. Upload a file first."];
    setStatus("Generation blocked.");
    showValidationMessages();
    return;
  }

  const selectedFile = getSelectedFile();
  if (!selectedFile) {
    state.errors = ["Select a student before generating."];
    setStatus("Generation blocked.");
    showValidationMessages();
    return;
  }

  const rows = state.groupedRows.get(selectedFile);
  if (!rows) {
    state.errors = [`Selected student "${selectedFile}" was not found in CSV data.`];
    setStatus("Generation blocked.");
    showValidationMessages();
    return;
  }

  if (rows.length !== 3) {
    state.errors = [`${selectedFile} has ${rows.length} review(s); requires exactly 3.`];
    setStatus("Generation blocked.");
    showValidationMessages();
    return;
  }

  onManualInputChange();
  const manual = ensureStudentInput(selectedFile);
  const mergedManual = {
    lsat: manual.lsat,
    gpa: manual.gpa,
    kjd: manual.kjd,
    urm: manual.urm,
    summary: manual.summary,
    nextSteps: manual.nextSteps,
    otherText: `${manual.kjd} • ${manual.urm}`,
  };

  const report = buildStudentReport(selectedFile, rows, mergedManual);
  state.currentReport = report;
  state.reports = [report];
  renderPreviewFromCurrent();
  downloadPdfBtn.disabled = false;
  setStatus(`Generated report for ${selectedFile}.`);
  showValidationMessages();
}

function onStudentSelectionChange() {
  syncManualFormFromSelection();
  state.currentReport = null;
  state.reports = [];
  downloadPdfBtn.disabled = true;
  previewRoot.innerHTML =
    '<p class="placeholder">Update metadata as needed, then click Generate Report.</p>';
}

function openPrintWindow(title, bodyHtml) {
  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    state.errors = ["Pop-up blocked. Allow pop-ups to print."];
    showValidationMessages();
    return;
  }

  const printHtml = `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <title>${escapeHtml(title)}</title>
    <style>${PRINT_CSS}</style>
  </head>
  <body>
    ${bodyHtml}
    <script>${PRINT_FIT_SCRIPT}</script>
  </body>
</html>`;

  printWindow.document.open();
  printWindow.document.write(printHtml);
  printWindow.document.close();
  printWindow.focus();

  try {
    setTimeout(() => {
      try {
        const bodyContent = printWindow.document?.body?.innerHTML?.trim() || "";
        if (!bodyContent) {
          state.warnings = [
            ...state.warnings,
            "Print window failed to render. Check popup settings and retry.",
          ];
          showValidationMessages();
          return;
        }
        printWindow.print();
      } catch (_error) {
        state.warnings = [
          ...state.warnings,
          "Print fallback encountered an issue. Retry printing if needed.",
        ];
        showValidationMessages();
      }
    }, 250);
  } catch (_error) {
    state.warnings = [
      ...state.warnings,
      "Could not trigger print fallback automatically. You can print manually from the popup.",
    ];
    showValidationMessages();
  }
}

function sanitizeFileName(value) {
  return String(value || "student-report")
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
}

async function onDownloadCurrentStudentPdf() {
  const report = state.currentReport;
  if (!report) return;

  const hasPdfLib = Boolean(window.jspdf?.jsPDF);
  const hasHtmlToImage = Boolean(window.htmlToImage?.toPng);
  const hasHtml2Canvas = typeof window.html2canvas === "function";

  if (!hasPdfLib || (!hasHtmlToImage && !hasHtml2Canvas)) {
    state.warnings = [
      ...state.warnings,
      "PDF library is unavailable. Use your browser Print dialog as fallback.",
    ];
    showValidationMessages();
    return;
  }

  const sourceDoc = previewRoot.querySelector(".doc-shell");
  if (!sourceDoc) {
    state.warnings = [
      ...state.warnings,
      "Could not find preview content. Generate the report again and retry PDF.",
    ];
    showValidationMessages();
    return;
  }

  fitAllSummarySections(previewRoot);
  await new Promise((resolve) => requestAnimationFrame(resolve));

  const staging = document.createElement("div");
  staging.style.position = "fixed";
  staging.style.left = "-10000px";
  staging.style.top = "0";
  staging.style.width = "612px";
  staging.style.opacity = "1";
  staging.style.pointerEvents = "none";
  staging.style.zIndex = "-1";
  staging.style.overflow = "hidden";

  const clone = sourceDoc.cloneNode(true);
  clone.classList.add("pdf-export-root");
  clone.style.width = "612px";
  clone.style.margin = "0";
  clone.style.padding = "0";
  staging.appendChild(clone);
  document.body.appendChild(staging);

  try {
    if (document.fonts?.ready) {
      await document.fonts.ready;
    }
    fitAllSummarySections(staging);
    paginateReaderCards(staging);
    paginateTagExplanationCards(staging);
    await new Promise((resolve) => requestAnimationFrame(resolve));
    fitAllSummarySections(staging);
    paginateReaderCards(staging);
    paginateTagExplanationCards(staging);
    await new Promise((resolve) => requestAnimationFrame(resolve));

    const filename = `${sanitizeFileName(report.fileName) || "student-report"}.pdf`;
    const pageNodes = [...clone.querySelectorAll(".page")];
    if (!pageNodes.length) {
      throw new Error("No pages found to export.");
    }
    const pdf = new window.jspdf.jsPDF({
      unit: "pt",
      format: [612, 792],
      orientation: "portrait",
    });

    for (let i = 0; i < pageNodes.length; i += 1) {
      let imageData = "";
      if (hasHtmlToImage) {
        imageData = await window.htmlToImage.toPng(pageNodes[i], {
          pixelRatio: 2,
          cacheBust: true,
          backgroundColor: "#ffffff",
          canvasWidth: 612,
          canvasHeight: 792,
          width: 612,
          height: 792,
        });
      } else {
        const canvas = await window.html2canvas(pageNodes[i], {
          scale: 2,
          useCORS: true,
          backgroundColor: "#ffffff",
          width: 612,
          height: 792,
          windowWidth: 612,
          windowHeight: 792,
          x: 0,
          y: 0,
          scrollX: 0,
          scrollY: 0,
        });
        imageData = canvas.toDataURL("image/jpeg", 0.98);
      }

      if (i > 0) {
        pdf.addPage([612, 792], "portrait");
      }
      pdf.addImage(imageData, "PNG", 0, 0, 612, 792, undefined, "FAST");
    }

    pdf.save(filename);
  } catch (_error) {
    state.warnings = [
      ...state.warnings,
      "PDF export failed. Use your browser Print dialog as fallback.",
    ];
    showValidationMessages();
  } finally {
    staging.remove();
  }
}

function onPrintCurrentStudent() {
  const report = state.currentReport;
  if (!report) return;
  openPrintWindow(report.fileName, `<div class="doc-shell">${renderStudentDocument(report)}</div>`);
}
