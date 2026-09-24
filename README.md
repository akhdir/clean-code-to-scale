# Clean Code to Scale

A 20-hour training curriculum (10 sessions × 2 hours) taking a software
engineering team from clean-code fundamentals through clean architecture,
refactoring, and coding with AI agents. Built for a mixed audience:
engineers from junior to senior, tech leads, QEs, and AI engineers.

Every session shares one running example — an order-processing service —
so the code you see in Day 4 is the same code that got refactored in
Day 1 and re-architected in Day 9.

## What's in this repo

- **`curriculum.html`** — the full course plan: objectives, format, and
  timing for all ten sessions.
- **`session-NN.html`** — facilitator and participant material for each
  built session: concept, live demo, hands-on kata, debrief, and
  facilitator prep notes with answer keys.
- **`code/day-NN-*/`** — the runnable code samples and kata starter files
  extracted out of each session, organized by day.

## Sessions built so far

| Day | Module | Session |
|---|---|---|
| 1 | Clean Code | Naming, Functions & Formatting |
| 2 | Clean Code | Error Handling & Code Smells |
| 3 | Clean Code | SOLID & Object Design |
| 4 | Refactoring | Refactoring Technique & the Smell Catalog |
| 9 | Clean Architecture | The Dependency Rule & Layered Boundaries |
| 10 | Clean Architecture | Ports, Adapters & a Real Case Study |

Days 5–8 (Refactoring Legacy Code, Scalability Fundamentals, Distributed
Systems & Resilience, Coding With AI Agents) are on the curriculum but
not yet built.

## Using the code

Each file under `code/` is a teaching snippet, not a package — most
reference illustrative stand-ins (`db`, `email_service`, `psycopg2`,
`SlackNotifier`, and so on) for real services, exactly as they appear in
the session material. `*_before.py` / `*_kata.py` files are the starting
point for an exercise; `*_after.py` / `*_solution.py` files are the
answer key.
