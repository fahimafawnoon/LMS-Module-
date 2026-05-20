# Course Builder Quick Reference Guide
> Use this alongside `LMS_MODULE_TEMPLATE.md` when building new courses.

---

## How to Build a New Course (Checklist)

```
STEP 1 — COPY THE TEMPLATE
  □ Duplicate LMS_MODULE_TEMPLATE.md
  □ Rename: courses/[COURSE-ID]_[Short-Title]/COURSE_DESIGN.md
  □ Assign a Course ID (e.g., SWV-002, OPS-001, SLS-003)

STEP 2 — FILL COURSE IDENTITY (Section 1)
  □ Write title using the naming formula
  □ Write tagline (1 sentence, second person)
  □ Write overview (75–100 words, lead with pain point)

STEP 3 — WRITE LEARNING OBJECTIVES (Section 2)
  □ Write 3–4 objectives
  □ Each starts with a Bloom's action verb
  □ Each maps to at least one module and one assessment question
  □ None use "understand," "know," or "be aware of"

STEP 4 — BUILD MODULE OUTLINES (Section 3)
  □ All 6 modules defined with duration, type, and content points
  □ Modules 1–5: at least one learning activity each
  □ Module 6: assessment questions ready (see Step 6)
  □ Total seat time target: 45–60 minutes

STEP 5 — DEFINE COMPLETION CRITERIA (Section 4)
  □ Every criterion is LMS-enforceable (not "learner should...")
  □ Passing score for assessment specified
  □ Failure path defined (how many attempts, what happens after)
  □ Timeline for completion from assignment date set

STEP 6 — WRITE ASSESSMENT QUESTIONS (Section 7)
  □ Minimum 15 questions in pool for 10-question assessment
  □ Every question tests Apply, Analyze, or Evaluate (not recall)
  □ Every question maps to a learning objective
  □ Every question has a rationale for the correct answer
  □ Distractors are plausible (not obviously wrong)

STEP 7 — COMPLETE METADATA (Section 8)
  □ All fields filled — especially short description (≤160 chars)
  □ Prerequisites listed or confirmed "None"
  □ Review cycle date set

STEP 8 — HANDOFF TO DEVELOPMENT
  □ COURSE_DESIGN.md reviewed by SME
  □ Video scripts drafted (separate document per module)
  □ Scenario scripts drafted (Module 4)
  □ Asset list created (slides, graphics, video shoots needed)
  □ LMS upload tested in staging environment
```

---

## Bloom's Taxonomy Quick Reference

| Level | Good Verbs | What it tests |
|-------|-----------|---------------|
| Remember | Define, List, Identify, Name, Recall | Basic recall |
| Understand | Describe, Explain, Summarize, Classify | Conceptual grasp |
| **Apply** | Use, Apply, Demonstrate, Solve, Execute | Transfer to new situation |
| **Analyze** | Distinguish, Compare, Examine, Differentiate | Breaking down complexity |
| **Evaluate** | Judge, Justify, Critique, Select, Prioritize | Making informed decisions |
| Create | Design, Construct, Develop, Formulate, Plan | Original production |

> Target Apply, Analyze, and Evaluate for most corporate learning objectives.
> Create is appropriate for leadership development and strategic planning courses.

---

## Common Mistakes — Avoid These

| Mistake | Better Approach |
|---------|----------------|
| Objectives that start with "understand" or "know" | Use Apply or Analyze verbs instead |
| Module durations listed as ranges ("10–15 min") | Pick a specific target (12 minutes) |
| Completion criteria that can't be LMS-enforced | Tie every criterion to a measurable LMS trigger |
| Assessment questions testing recall ("What does X stand for?") | Test application ("A client does X — what do you do?") |
| Action plan that's a survey (rating scales only) | Include free-text commitments, specific dates, accountability partners |
| Module 6 as passive review | Make it a synthesis activity + forward-facing close |
| Generic scenario characters | Name the characters, give them context, make stakes real |

---

## File Naming Convention

```
Repository structure:
/LMS-Module-
  LMS_MODULE_TEMPLATE.md          ← Master reusable template
  COURSE_BUILDER_GUIDE.md         ← This file
  /courses
    /[ID]_[Short-Title]/
      COURSE_DESIGN.md            ← Filled-in template for this course
      /scripts/                   ← Video + scenario scripts (one file per module)
      /assets/                    ← Slide decks, graphics, handouts
      /assessment/                ← Full question bank (if separate from design doc)
```

---

## Module Type Selector

Use this to choose the right content format for each module:

| Module Purpose | Best Format |
|---------------|-------------|
| Orientation / Why this matters | Short video (leadership voice) |
| Core concept / mental model | Narrated slides or video + slide deck |
| Demonstrate correct behavior | Contrast scenario video (wrong vs. right) |
| Practice / application | Branching scenario (Storyline or Rise) |
| Reflection / transfer | Reading + in-LMS form |
| Assessment / capstone | Graded quiz + certificate |

---

*Quick reference for SWV course series and all future LMS builds.*
