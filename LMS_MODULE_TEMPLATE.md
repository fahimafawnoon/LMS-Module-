# LMS Module Template — Reusable Framework
> **Version:** 1.0 | **Standard:** LinkedIn Learning Compatible
> **Usage:** Duplicate this file for each new course. Replace bracketed placeholders with course-specific content. Keep all structural elements intact.

---

## TEMPLATE SECTION 1: COURSE IDENTITY

### Course Title
`[Action Verb Phrase]: [Specific Skill or Mindset] for [Target Audience]`

**Naming Rules:**
- Lead with a verb or outcome-signal word (e.g., "Mastering," "Building," "The [Brand] Way")
- Keep it under 10 words
- Avoid buzzwords: no "leveraging," "synergizing," or "next-level"

**Example:**
> "The Swivel Way: Premium Mindset & Client-Ready Communication"

---

### Course Tagline (1 sentence)
`[Role or persona] who [completes this course] will [specific, observable outcome] in [timeframe or context].`

---

### Course Overview (75–100 words)
Write in second person ("you"), present tense. Lead with the problem or gap, follow with the outcome, close with a credibility signal or stakes statement.

```
[PAIN POINT OR CONTEXT SENTENCE]. This course gives you [specific tools/frameworks/mindsets]
to [measurable outcome]. By the end, you will [concrete capability #1] and [concrete capability #2]—
skills that apply [immediately / on your next call / in your next client interaction].
[Optional: credibility line about why this matters at your organization / in this industry.]
```

---

## TEMPLATE SECTION 2: LEARNING OBJECTIVES

**Rules:**
- Write exactly 3–4 objectives
- Lead each with a Bloom's Taxonomy action verb (not "understand" or "know")
- Each objective must be independently verifiable in assessment
- Map each objective to at least one module and one assessment question

| # | Objective | Bloom's Level | Maps to Module | Assessed In |
|---|-----------|--------------|----------------|-------------|
| 1 | [Verb] + [specific skill/behavior] + [context/condition] | [Remember/Apply/Analyze/Evaluate] | Module [X] | Quiz Q[X] |
| 2 | [Verb] + [specific skill/behavior] + [context/condition] | [Apply/Analyze] | Module [X] | Scenario / Reflection |
| 3 | [Verb] + [specific skill/behavior] + [context/condition] | [Apply/Evaluate] | Module [X] | Quiz Q[X] |
| 4 | [Verb] + [specific skill/behavior] + [context/condition] | [Evaluate/Create] | Module [X] | Action Plan |

---

## TEMPLATE SECTION 3: MODULE BREAKDOWN

**Total Course Target:** 45–60 minutes seat time
**Module Count:** 6 modules (5 content + 1 assessment/capstone)

---

### Module Structure Schema

For each of the 6 modules, complete this block:

```
MODULE [N]: [Title]
─────────────────────────────────────────────────────────────
Duration:     [X] minutes
Content Type: [Video | Reading | Scenario | Reflection | Assessment]
Format:       [Single video / Branching scenario / Slide deck + narration / Text + visuals]
Purpose:      [What this module accomplishes in the learner's journey — 1 sentence]

CONTENT OUTLINE:
  1. [Key point or segment name] — [~time or word count]
  2. [Key point or segment name] — [~time or word count]
  3. [Key point or segment name] — [~time or word count]
  [Add rows as needed — target 3–5 content points per module]

LEARNING ACTIVITY:
  Type:     [Knowledge check / Reflection prompt / Scenario decision / None]
  Prompt:   "[Exact text of the activity prompt or question]"
  Purpose:  [What this activity reinforces]

COMPLETION TRIGGER:
  [Video 100% watched | Quiz passed at X% | Reflection submitted | Time-on-page ≥ X min]
─────────────────────────────────────────────────────────────
```

---

### Module 1 — Orientation / Why This Matters
- **Role:** Sets stakes, establishes relevance, introduces the course framework
- **Tone:** Direct, motivating — answer "what's in it for me?" immediately
- **Avoid:** Long introductions, corporate history, excessive context-setting

### Module 2 — Core Concept / Mindset Foundation
- **Role:** Introduces the primary mental model or framework learners will use
- **Tone:** Instructional but grounded — use a real analogy or case example
- **Avoid:** Abstract theory without a concrete anchor

### Module 3 — Skill Demonstration / Model the Behavior
- **Role:** Shows the concept in action — video scenario or annotated example
- **Tone:** Observational — learner watches, identifies, and names what's happening
- **Avoid:** Telling learners what they just saw; let them articulate it

### Module 4 — Skill Practice / Application
- **Role:** Learner applies the concept in a controlled scenario
- **Tone:** Interactive, low-stakes — frame mistakes as data, not failure
- **Avoid:** Single right-answer scenarios for nuanced skills

### Module 5 — Integration / Real-World Transfer
- **Role:** Connects course content to the learner's actual role and context
- **Tone:** Reflective — use "in your role" and "on your next [interaction]"
- **Avoid:** Generic workplace examples; be specific to the target role

### Module 6 — Assessment + Action Plan (Capstone)
- **Role:** Verifies learning, creates accountability, closes the loop
- **Tone:** Confident, forward-facing — "you now have the tools to..."
- **Avoid:** Passive review; make the learner synthesize and commit

---

## TEMPLATE SECTION 4: COMPLETION CRITERIA

Completion must be **specific** and **technically enforceable** in your LMS.

```
COURSE COMPLETION REQUIREMENTS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
□  All 6 modules opened and marked complete
□  Minimum time-on-course: [X] minutes (LMS-tracked)
□  Knowledge Assessment: score ≥ [X]% (e.g., 80%)
   — Learners may retake [X] times before manager notification
   — On failure: directed to review [specific module(s)]
□  Reflection/Action Plan: submitted (not graded for correctness)
□  Course survey: submitted (optional / required — choose one)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Completion Status in LMS: "Complete" unlocks after ALL boxes above are met.
Incomplete after [X days from assignment]: triggers manager alert / auto-reminder.
```

---

## TEMPLATE SECTION 5: LMS TECHNICAL SPECIFICATIONS

### Video
| Spec | Requirement |
|------|-------------|
| Resolution | 1920×1080 (1080p minimum) |
| Aspect Ratio | 16:9 |
| Frame Rate | 30fps |
| Audio | Stereo, -14 LUFS, no background noise |
| Closed Captions | Required — SRT file delivered with video |
| Max File Size | 500MB per video file |
| Hosting Format | MP4 (H.264 codec) |
| Accessibility | Captions + audio description for visual-only content |

### Slides / Screen Content
| Spec | Requirement |
|------|-------------|
| Format | 16:9 PowerPoint or Google Slides source file |
| Fonts | [Brand font stack] — minimum 24pt body, 36pt headers |
| Colors | Brand palette only — minimum 4.5:1 contrast ratio (WCAG AA) |
| Export | PDF (learner download) + PNG sequence for video integration |
| Slide Count | Target ≤ 8 slides per module (avoid slide decks as primary content) |

### Quiz / Assessment
| Spec | Requirement |
|------|-------------|
| Question Types | Multiple choice (single answer), multiple select, scenario-based |
| Questions per Module | 1–2 knowledge checks (ungraded) per content module |
| Final Assessment | 5–10 questions, graded, passing score = [X]% |
| Randomization | Question pool of 1.5× final count (15 questions for 10-question test) |
| Feedback | Immediate per-question feedback — correct answer + 1-sentence rationale |
| Attempts Allowed | [X] attempts — define escalation path on repeat failure |

### SCORM / xAPI
| Spec | Requirement |
|------|-------------|
| Standard | SCORM 1.2 or xAPI (Tin Can) — confirm with LMS vendor |
| Completion Trigger | `cmi.core.lesson_status = passed` OR xAPI `result.success = true` |
| Score Reporting | Pass/fail + raw score transmitted to LMS gradebook |
| Suspend/Resume | Required — learners must be able to exit and re-enter |
| Mobile | Responsive — tested on iOS Safari and Android Chrome |

### Accessibility
- All videos: closed captions (SRT) + transcript (TXT or PDF)
- All images: alt text in authoring tool
- All interactive elements: keyboard-navigable
- Color: never as sole conveyor of meaning
- Compliance Target: WCAG 2.1 Level AA

---

## TEMPLATE SECTION 6: REFLECTION & ACTION PLAN TEMPLATE

> This is the **exact template** learners complete in Module 5 or 6.
> Deliver it as an in-LMS text entry, a downloadable PDF form, or both.

---

```
╔══════════════════════════════════════════════════════════════╗
║          [COURSE NAME] — MY ACTION PLAN                      ║
║          Learner: ________________  Date: ___________        ║
╚══════════════════════════════════════════════════════════════╝

PART 1: WHAT I LEARNED (Synthesis)
────────────────────────────────────────────────────────────────
The most important thing I'm taking away from this course is:

[3–5 sentence free response]


PART 2: WHERE I AM NOW (Honest Assessment)
────────────────────────────────────────────────────────────────
Rate yourself honestly on each course skill (1 = not yet doing this,
5 = doing this consistently and well):

  [ ] [Skill/Behavior #1 from learning objectives] — Rating: ___/5
  [ ] [Skill/Behavior #2 from learning objectives] — Rating: ___/5
  [ ] [Skill/Behavior #3 from learning objectives] — Rating: ___/5

The area where I have the most room to grow is:
[1–2 sentences]


PART 3: MY COMMITMENTS (Specific Actions)
────────────────────────────────────────────────────────────────
In the next 7 days, I will practice [specific behavior] by:
  Action:    ___________________________________________________
  Context:   When/where this will happen: _____________________
  Evidence:  How I'll know I did it: __________________________

In the next 30 days, I will achieve [measurable outcome] by:
  Action:    ___________________________________________________
  I'll ask   [name/role] to hold me accountable.


PART 4: ONE THING I'LL STOP / START / CONTINUE
────────────────────────────────────────────────────────────────
  STOP:     ___________________________________________________
  START:    ___________________________________________________
  CONTINUE: ___________________________________________________


PART 5: SHARE WITH YOUR MANAGER (Optional but Recommended)
────────────────────────────────────────────────────────────────
One thing from this course I want to discuss with my manager:

[Free response]
```

---

## TEMPLATE SECTION 7: SAMPLE ASSESSMENT QUESTIONS

> Write 5–10 questions per course. Use this schema for each question.
> Questions should test **application and analysis**, not recall.

---

**Question Schema:**
```
Q[N]. [Scenario setup if needed. Then the question.]

  A) [Plausible but incorrect — common misconception]
  B) [Correct answer — specific and unambiguous]
  C) [Plausible but incomplete — missing a key element]
  D) [Plausible but incorrect — opposite or extreme position]

Correct Answer: B
Rationale: [1–2 sentences explaining why B is correct and why the others aren't.]
Objective Mapped: [Learning Objective #]
Bloom's Level: [Apply / Analyze / Evaluate]
```

---

## TEMPLATE SECTION 8: COURSE METADATA (LMS Upload Fields)

```
Course ID:          [AUTO or MANUAL — e.g., SWV-001]
Course Title:       [Full title]
Short Description:  [≤160 characters — used in catalog previews]
Long Description:   [Full overview paragraph]
Duration (minutes): [Exact seat-time estimate]
Difficulty Level:   [Beginner / Intermediate / Advanced]
Target Audience:    [Job role(s) or personas]
Prerequisites:      [None / List specific courses]
Category/Tags:      [Taxonomy tags for LMS catalog search]
Author/SME:         [Name, title]
Last Reviewed:      [Date — set review cycle, e.g., annually]
Language:           [English (US) / Other]
Completion Type:    [Pass/Fail with score | Completion only]
Passing Score:      [X%]
Certificate:        [Yes / No — if yes, specify template]
```

---

*End of Reusable Template — All sections above are duplicable for any new course.*
*Replace all bracketed placeholders. Do not remove structural elements.*
