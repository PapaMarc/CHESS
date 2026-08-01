# CHESS Modernization Roadmap

## Purpose

This document is the long-term modernization reference for the legacy CHESS project. It is written for:

- Humans revisiting architecture and execution decisions.
- A Copilot assistant that needs immediate context before starting implementation.

It captures:

- What the old system does and how it is organized.
- What to keep, what to replace, and how to modularize the kept logic.
- The target architecture: .NET server plus SQL-backed state.
- How to support multiple clients from one authoritative backend.
- Risks, sequencing, and decision checkpoints.

---

## 1) Legacy System Summary

### 1.1 What this project is

The current codebase is a legacy Visual Basic 4-era chess application designed around Exchange/MAPI custom form workflows. Gameplay is turn-based and message-driven:

- A player opens or composes a message.
- The chess UI is loaded through MAPI form hosting.
- A move is made.
- A new message is sent with updated game state and metadata.

### 1.2 Main files and responsibilities

- CHESS.VBP: project references and module/form wiring.
- CHSMAIN.BAS: startup, MAPI object/session flow, message open/compose/reply, send logic, stuffing/unstuffing game payload fields.
- GAMEUTIL.BAS: chess rules and board logic (move validation, check/checkmate, castling, en passant, replay logic, notation assembly, board orientation helpers).
- GAME.FRM (+ FRX): main board UI and interaction flow (drag/drop pieces, move list, comments, send/undo/surrender behavior).
- NEWGAME.FRM (+ FRX): new match setup and opponent selection.
- PAWNEXCH.FRM: pawn promotion chooser.
- FRMADMIN.FRM: hidden admin/mapi host form and lifecycle bridge.
- GUID.BAS: conversation index/GUID/time helper logic for message threading metadata.
- TASKLIST.BAS: shell/window/tasklist-related plumbing.
- CONSTANT.BAS: misc constants.

### 1.3 Core behavioral model

- Stateful global variables represent current game and session state.
- Message fields carry serialized game data and move metadata.
- UI events and game rules are tightly coupled.
- Transport and persistence are effectively embedded in email/MAPI semantics.

### 1.4 Legacy strengths worth preserving

- Mature chess-rule behavior.
- Practical move replay model.
- User-facing move history and comments.
- Message-per-move mental model (good conceptual bridge to API events).

### 1.5 Legacy constraints that block direct porting

- Obsolete Exchange/MAPI custom form hosting model.
- Tight coupling between rules and UI controls.
- Global mutable state across modules.
- Legacy platform calls and assumptions.

---

## 2) Modernization Goals

### 2.1 Primary goals

- Make backend authoritative for game validity, turn order, and persisted state.
- Separate domain logic from UI and transport concerns.
- Support multiple clients (web, desktop, mobile) from one API.
- Preserve legacy game behavior where feasible.
- Keep the path open for incremental feature growth.

### 2.2 Non-goals (initial phase)

- Exact recreation of Exchange custom form hosting.
- Pixel-perfect clone of legacy VB UI.
- Full online tournament ecosystem in phase 1.
- Importing legacy Exchange custom-form chess games.

---

## 3) Keep vs Replace

### 3.1 Keep (conceptual and code-port targets)

Keep and port the chess domain behavior from GAMEUTIL.BAS into a new pure domain engine:

- Board model and piece model.
- Legal move validation rules.
- Special moves: castling, en passant, promotion.
- Check and checkmate detection.
- Move replay from sequence.
- Move notation generation and history support.

### 3.2 Replace (full rewrite)

Replace all Exchange/MAPI and VB forms integration:

- Message/session transport flow in CHSMAIN.BAS.
- UI forms and event coupling (GAME/NEWGAME/PAWNEXCH forms).
- MAPI host/admin bootstrapping.
- Legacy shell/window/tasklist glue.

### 3.3 Preserve where useful as compatibility adapters

- Legacy payload parsing ideas can inspire diagnostics and data-model design (not import tooling in MVP).
- Existing move history semantics can map to modern SAN/PGN representations.

---

## 4) Target Architecture

### 4.1 High-level shape

A server-authoritative architecture:

- Backend: ASP.NET Core Web API.
- Database: SQL (start SQLite for local development if needed, target PostgreSQL or SQL Server for production).
- Auth: multi-provider identity baseline including GitHub, Microsoft, Google, and email/password.
- Clients: browser, Windows desktop, Android (and optionally iOS).

### 4.2 Bounded contexts/modules

Define explicit modules in the new solution:

1. Domain Engine

- Pure chess model and rules.
- No database, no HTTP, no UI dependencies.

2. Application Layer

- Use cases: create game, join game, submit move, resign, request draw, load replay state.
- Validation and orchestration.

3. Infrastructure Layer

- Persistence via EF Core or equivalent.
- Auth integration.
- Messaging/notification adapters.

4. API Layer

- REST endpoints plus optional real-time channel (SignalR/WebSocket).
- DTO mapping and request/response contracts.

5. Client Apps

- Render board and move list.
- Call API.
- Maintain local view state only.

### 4.3 Server-authoritative rules (must-have)

- The server validates every move.
- The server enforces whose turn it is.
- The server is source of truth for game state and status.
- Clients cannot directly assert new board states.

---

## 5) Data Model (SQL)

Recommended baseline tables:

1. Users

- Id
- ExternalProvider
- ExternalSubjectId
- DisplayName
- Email
- CreatedUtc

2. Games

- Id (GUID)
- WhiteUserId
- BlackUserId
- Status (Active, Checkmate, Resigned, Drawn, Abandoned)
- SideToMove
- CurrentPly
- CurrentFen (or equivalent canonical position representation)
- WinnerUserId (nullable)
- CreatedUtc
- UpdatedUtc
- ConcurrencyToken (row version)

3. Moves

- Id
- GameId
- Ply
- PlayedByUserId
- SAN
- UCI (optional but useful)
- FromSquare
- ToSquare
- PromotionPiece (nullable)
- ResultingFen
- CreatedUtc

4. GameSnapshots (optional in phase 1, recommended in phase 2)

- Id
- GameId
- Ply
- Fen
- CreatedUtc

5. GameEvents (optional)

- Id
- GameId
- EventType
- PayloadJson
- CreatedUtc

Notes:

- Keep an index on (GameId, Ply).
- Enforce unique (GameId, Ply).
- Use optimistic concurrency at game row level.

---

## 6) Notation, Replay, and Board Navigation

### 6.1 Standards to adopt

- SAN for move display.
- PGN export/import for interoperability.
- FEN (or equivalent canonical representation) for per-position state.

### 6.2 Replay UX requirements

- Display move list under board.
- Click any move to jump to that state.
- Next/previous move stepping.
- Return-to-live action when browsing history.

### 6.3 Replay implementation pattern

- Store each move plus resulting position.
- Optionally snapshot every N plies for fast random access.
- Client asks server for state at ply N or reconstructs from returned move list.

---

## 7) API Contract Shape (Reference)

Suggested endpoints:

- POST /api/games
- GET /api/games/{gameId}
- GET /api/games/{gameId}/moves
- POST /api/games/{gameId}/moves
- POST /api/games/{gameId}/resign
- POST /api/games/{gameId}/draw-offer
- POST /api/games/{gameId}/draw-accept
- GET /api/games/{gameId}/state?ply={n}
- GET /api/games/{gameId}/pgn

Move submission request should include:

- expectedPly or concurrency token
- from/to squares
- optional promotion piece

Server response should include:

- accepted move details (SAN/UCI)
- resulting position
- updated status
- next side to move

---

## 8) Multi-Client Strategy

### 8.1 Principle

One backend, many clients. Domain rules remain centralized on server.

### 8.2 Client responsibilities

- Render board and move history.
- Handle user interaction.
- Call API and display authoritative responses.
- Cache local view state for UX.

### 8.3 Supported client suite path

1. Web client first (fastest path to feedback).
2. Windows desktop client next (if needed for richer native UX).
3. Android client once API stabilizes.

### 8.4 Shared client concerns

- Authentication token handling.
- Network retry and offline behavior.
- State refresh after reconnect.
- Replay and move list consistency.

---

## 9) Implementation Pattern and Sequencing

### Phase 0: Discovery and baseline tests

- Inventory legacy rules and edge cases.
- Build test vectors from known move scenarios (including special moves).
- Decide canonical board/state representation.

### Phase 1: Domain engine extraction

- Implement chess engine as pure .NET class library.
- Add comprehensive unit tests (rules, check/checkmate, special moves, notation).
- Ensure deterministic state transitions.

### Phase 2: Persistence and API

- Introduce SQL schema and migrations.
- Implement core game and move endpoints.
- Add integration tests for move transaction flow.

### Phase 3: Web client MVP

- Board rendering.
- Move submission.
- Move list and replay navigation.
- Game status and result handling.

### Phase 4: Multi-client enablement

- Harden auth and API contracts.
- Add real-time updates.
- Add desktop/mobile clients.

### Phase 5: Compatibility and interoperability enhancements (optional)

- Add PGN export/import.
- Add external game import from standardized formats/systems as a future option.
- Keep legacy Exchange-game import explicitly out of scope unless requirements change.

---

## 10) Things to Keep in Mind (Design and Risk)

1. Coupling risk

- Avoid embedding UI assumptions in domain logic.
- Domain engine must remain transport-agnostic.

2. Concurrency risk

- Two clients may submit nearly simultaneously.
- Use optimistic concurrency and reject stale move submissions.

3. Trust boundary

- Never trust client-submitted board state as authoritative.
- Validate move legality server-side every time.

4. Identity and authorization

- Ensure only game participants can view/play the game.
- Check turn ownership on submit.

5. Replay correctness

- Move indexing must be unambiguous (ply vs move number clarity).
- Keep stable ordering and idempotency rules.

6. Interop and portability

- Favor standards (SAN, PGN, FEN) to reduce lock-in.

7. Observability

- Add structured logs for move validation failures and state transitions.

8. Product decisions to settle early

- Real-time requirement or turn-based polling in v1.
- Draw handling and timeout rules.
- Whether comments/annotations are first-class in v1.

---

## 11) Suggested Solution Layout

- src/Chess.Domain
- src/Chess.Application
- src/Chess.Infrastructure
- src/Chess.Api
- src/Chess.WebClient
- src/Chess.Contracts
- tests/Chess.Domain.Tests
- tests/Chess.Api.IntegrationTests

---

## 12) Definition of Done for Initial Port

Initial modernization milestone is done when:

1. A full legal game can be played through API and persisted in SQL.
2. Move list is shown in SAN and replay is navigable move-by-move.
3. Clicking any move renders corresponding board state.
4. Server enforces turn and legality.
5. At least one client (web) is production-usable.
6. Core domain tests cover critical rules and regressions.

---

## 13) Copilot Handoff Starter (for next implementation session)

Use this as the first prompt context for a coding assistant:

- Goal: Begin phase 1 domain extraction from legacy VB rules into a pure .NET library.
- Priorities:
  1. Define board state and move models.
  2. Implement legal move validation and special moves.
  3. Add tests mirroring legacy behavior.
  4. Keep domain engine independent of persistence/UI.
- Constraints:
  - Server-authoritative architecture.
  - SAN plus replay support required.
  - Design for multiple clients from one API.

Suggested first coding tasks:

1. Create Chess.Domain with core entities and value objects.
2. Add MoveApplier service and legal move validator.
3. Add FEN serializer/deserializer.
4. Add SAN formatter.
5. Add tests for:

- Castling legality and outcomes.
- En passant eligibility and execution.
- Promotion handling.
- Check/checkmate detection.
- Replay consistency.

---

## 14) Open Questions to Resolve Before Full Build

1. Production SQL choice: PostgreSQL or SQL Server?
2. Authentication baseline and rollout: GitHub, Microsoft, Google, and email/password. Which providers are required at MVP vs phase 2?
3. Do we need correspondence-email notifications in v1?
4. Is draw-offer workflow required in MVP?
5. Should comments/annotations per move be included in MVP?
6. External standardized game import (for example PGN from other systems): is this required in MVP or deferred?
7. Confirmed scope boundary: legacy Exchange custom-form game import is not required.

---

## 15) Decision Log Template

Keep this section updated as decisions are made.

- Date:
- Decision:
- Context:
- Alternatives considered:
- Chosen option:
- Consequences:

---

## 16) Summary

Modernization is best approached as a controlled rewrite around a preserved domain core:

- Preserve chess behavior.
- Replace legacy hosting/transport/UI architecture.
- Move to a .NET server with SQL authoritative state.
- Expose stable APIs for web, desktop, and mobile clients.
- Build replay and notation as first-class capabilities.

This roadmap is the reference baseline for kickoff and future continuation.
