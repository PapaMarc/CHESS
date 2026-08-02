# CHESS Modernization Roadmap B-WebClient

## Purpose

This document is a client-variation of Roadmap B.

Roadmap B defines the underlying store-and-forward, Microsoft Graph-based game model. This roadmap keeps that model intact, but changes the first client implementation to a web application rather than a native Windows client.

Primary intent:

- Build the web client first.
- Use Microsoft account sign-in and Microsoft Graph as the message transport and persistence layer.
- Preserve the legacy message-per-move mental model.
- Make each successive message self-contained with full current game state plus growing move history.
- Support opening any historical message to inspect and replay prior states.
- Treat this one client implementation as the practical path for all supported platforms in Roadmap B.

This document is the architecture and execution reference for implementing that web-first variant.

---

## 1) Relationship to Roadmap B

### 1.1 What stays the same

- Microsoft Graph remains the transport and persistence mechanism.
- Game state remains message-based and self-contained.
- Core gameplay logic remains in shared C# libraries, including legal move validation, special moves, check detection, checkmate detection, and replay/state reconstruction.
- Replay is message-centric and history aware.
- Chain validation, conflict handling, and head selection remain as defined in Roadmap B.
- The domain engine remains portable and UI-agnostic.

### 1.2 What changes

- The first client is a browser-based web application.
- Authentication is optimized for Microsoft account sign-in through browser flows.
- The UX is designed to work well with both desktop and mobile-sized layouts.
- The client should feel like one implementation that can serve the effective needs of the B roadmap across form factors.

### 1.3 Why this is still Roadmap B

- The core architecture remains store-and-forward via Graph.
- The move/message semantics remain unchanged.
- The difference is client surface, not game authority or the core chess rules engine.

---

## 2) Core Client Goals

1. Web-first access

- Users should be able to play directly in a browser.
- The client should be usable without a desktop install.

2. Responsive UI

- The board, controls, and message composition panes should adapt to available width.
- On smaller screens, the layout should compress into a mobile-friendly presentation.
- On larger screens, the same UI can expand to show richer board, history, and message details.

3. Touch and drag interaction

- Pieces should be movable with mouse or touch.
- The interaction model should feel natural on both pointer and touch devices.

4. Explicit send boundary

- The user can manipulate the local board state while preparing a move.
- A separate `Send Move` action advances the game by posting the next Graph message.

5. Invite flow

- The client should support entering and sending the initial email invite.
- A most-recently-used recipient list should reduce repeated entry work.
- Individual MRU entries should be removable.

---

## 3) High-Level Architecture

### 3.1 Client shape

- Browser-based SPA built with Blazor WebAssembly.
- Shared domain/application logic reused from the B roadmap in C# class libraries.
- UI-specific state is local to the web client.

### 3.2 Hosting model

- The browser client can be hosted as static content on GitHub Pages.
- Any Graph auth flow must remain compatible with a public-client browser app.
- The client should not require a dedicated server just to render or play the game.

### 3.3 Platform coverage assumption

- This single web client is the practical implementation target.
- Because it runs in the browser, it implicitly covers desktop and mobile access paths without separate native client code.

### 3.4 Identity and auth

- Microsoft account sign-in via browser-based MSAL flows.
- Graph scopes should remain minimal and aligned with mail/game-folder operations.

### 3.5 Shared code reuse

- The reusable chess and message-handling logic should remain in C# libraries.
- That keeps the domain portable for a future Windows or .NET MAUI client if one is later desired.
- The browser UI is the presentation layer, not the place where the core chess rules should live.

---

## 4) Responsive Layout Model

### 4.1 Design principle

- The same client should scale from narrow to wide screens.
- The layout should not be a fixed desktop-only board.

### 4.2 Behavior targets

- On wide screens, the board can sit beside move history and message controls.
- On medium screens, panels may stack or collapse.
- On narrow screens, the board should remain usable with touch, reduced chrome, and compact controls.

### 4.3 Threshold behavior

- The client should support breakpoint-driven layout changes similar to browser responsive tooling.
- Below a configurable width threshold, the UI can switch to a compact mobile-friendly arrangement.
- Above that threshold, it can expand into a fuller board-and-sidebar layout.

### 4.4 Practical consequence

- The web client should be developed with responsive behavior from the start, not retrofitted later.

---

## 5) Invitation and Recipient UX

### 5.1 Initial email invite

- Provide a UI for composing the initial game invite.
- The invite should integrate with the same Graph transport model used for game moves.

### 5.2 MRU recipient list

- Keep a list of recently used email addresses.
- Reuse the MRU list to speed up invite entry.
- Allow deleting individual MRU entries.

### 5.3 Storage approach

- The MRU list should be local to the client unless there is a later requirement to roam it.
- It should be treated as convenience data, not authoritative game state.

---

## 6) Board Interaction and Replay

### 6.1 Board controls

- The board should support drag and drop.
- The board should also support touch-driven movement on mobile and tablet-sized screens.
- Move entry should use drag-and-release as the default interaction.
- Releasing a piece on an invalid destination should snap the piece back to its origin square.

### 6.1a Board overlays and toggles

- The board chrome should expose independent hint toggles:
- Path: Destinations checkbox.
- Path: Telegraph checkbox.
- Notation checkbox.
- Path: Destinations controls whether legal destination squares are shown while a piece is selected or being dragged.
- Path: Telegraph controls whether path overlays are rendered during drag toward candidate destinations.
- Notation controls whether board labels (a-h and 1-8) are shown.
- Both Path toggles may be disabled at the same time.
- These toggles affect rendering only and must not alter move legality rules.

### 6.2 Replay navigation

- Provide forward and back buttons for navigating game/message history.
- The control set should adapt to the current view mode and the selected message state.

### 6.3 Historical inspection

- Opening an old message should render the board at that state.
- Replay controls should allow stepping through the move history from that point.
- A return-to-current action should restore the current head view.

### 6.4 Send Move boundary

- The user may make a local move on the board.
- The move is not committed until `Send Move` is pressed.
- Pressing `Send Move` creates the next Graph message and passes the turn onward.

### 6.5 Undo semantics

- Undo is supported only before `Send Move` is pressed.
- An undo button or icon can appear next to `Send Move` while a move is still local and uncommitted.
- After `Send Move`, undo is not supported.
- When no local move is pending, the undo control should be disabled or grayed out.
- This is acceptable because the message-based model does not persist the new state until send time, so there is no meaningful committed state to reverse before send, and no need to support a faux "pro" mode that hides or complicates undo.

---

## 7) Graph and Message Model

### 7.1 Unchanged transport model

- Continue using the Roadmap B message chain model.
- Each move message remains self-contained.
- The client derives state from the chain, not from a server-authoritative session.

### 7.2 Send flow

- User signs in.
- User creates or opens a game thread.
- User composes or selects an invite recipient.
- User makes a move in the board UI.
- User explicitly sends the move as a Graph message.

### 7.3 Validation

- The same chain validation and move legality checks from Roadmap B still apply.
- The web client must not assume its local board state is authoritative.

---

## 8) Implementation Plan

### Phase 0: Shape confirmation

- Confirm that this is the web-client variation of Roadmap B.
- Confirm the responsive layout target and breakpoint behavior.
- Confirm invite MRU storage scope.

### Phase 1: Shared core extraction

- Port or reuse the domain and Graph sync logic as reusable layers.
- Keep these layers independent of browser UI concerns.

### Phase 2: Web client MVP

- MS account sign-in.
- Invite composer with MRU list.
- Board rendering.
- Move interaction with touch and drag.
- Drag-release invalid-drop snap-back behavior.
- Board overlay toggles for Path: Destinations, Path: Telegraph, and Notation.
- Send Move action.

### Phase 3: Replay and history

- Back and forward navigation.
- Historical message inspection.
- Return to current head.

### Phase 4: Responsive polish

- Compact layout for smaller screens.
- Expanded layout for larger screens.
- Layout validation at multiple viewport sizes.

### Phase 5: Hardening

- Error handling for Graph failures and stale state.
- Conflict handling and branch surfacing.
- Persistence of MRU convenience data.

---

## 9) Risks and Mitigations

1. Responsive UI complexity

- Mitigation: design the layout as responsive from the start.

2. Touch and drag behavior differences

- Mitigation: test pointer and touch interactions early, especially on smaller screens.

3. User confusion between local board changes and committed moves

- Mitigation: keep `Send Move` visually distinct from board manipulation.

4. Graph latency or auth friction

- Mitigation: use clear loading and retry states in the web client.

5. MRU hygiene

- Mitigation: make per-entry removal simple and visible.

---

## 10) Definition of Done for Roadmap B-WebClient MVP

MVP is complete when:

1. A user can sign in with Microsoft account in the browser.
2. A user can enter an initial invite recipient, reuse MRU addresses, and delete individual MRU entries.
3. A user can view and play the chess board in a browser.
4. A user can drag or touch-move pieces.
5. A user can step backward and forward through the message/game state.
6. A user can press `Send Move` to commit the next move as a Graph message.
7. The UI remains usable across narrow and wide screen layouts.
