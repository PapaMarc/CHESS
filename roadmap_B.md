# CHESS Modernization Roadmap B

## Purpose

This roadmap is a variation fork of Roadmap A. It defines a serverless, store-and-forward architecture that does not require a hosted game server.

Primary intent:

- Build a Windows application first.
- Use Microsoft Graph as message transport and persistence.
- Keep legacy message-per-move mental model.
- Make each successive message self-contained with full current game state plus growing move history.
- Support opening any historical message to inspect and replay prior states.

This document is the architecture and execution reference for implementing that model.

---

## 1) Architecture Summary

### 1.1 High-level model

- A game is an append-only message chain.
- Every move creates a new message that references a parent message.
- Every message contains a complete game snapshot and replay metadata.
- Clients derive current state by validating and selecting the best head in the chain.

There is no central authoritative backend. Authority is derived from deterministic chain validation rules.

### 1.2 Platform target

Phase 1 client:

- Windows app (recommended .NET 9 + WinUI 3 or .NET MAUI on Windows target).

Future clients:

- Android and iPhone via shared domain/application layers.

### 1.3 Why this matches legacy behavior

- Legacy app used transport messages as durable game checkpoints.
- This design modernizes that model using Microsoft Graph and explicit schemas.

---

## 2) Core Design Principles

1. Deterministic chain validity

- Any client should derive the same valid game head given the same message set.

2. Self-contained messages

- Each move message can reconstruct board state without fetching all prior payloads.

3. Replay-first UX

- Historical message opens board at that message state with full step navigation.

4. Explicit conflict handling

- Concurrent or branched moves are detected and resolved by policy, not hidden.

5. Portable domain engine

- Chess rules remain pure and reusable across Windows/Android/iPhone.

---

## 3) Non-goals (initial)

- No hosted API for authority, matchmaking, or websocket updates.
- No background server-managed merge of conflicting branches.
- No import of legacy Exchange custom forms in phase 1.

---

## 4) Microsoft Graph Storage and Transport Model

### 4.1 Message container strategy

Recommended:

- Dedicated mail folder per game under an app-managed root folder, for example:
  - CHESS
  - CHESS/{GameId}

Each move is sent as a normal mail message to the opponent and copied into the game folder by sender and recipient clients.

Alternative (acceptable):

- Single folder with query by `GameId` extension property.

### 4.2 Message format location

Preferred payload location:

- JSON attachment (`application/json`) named `game-state.json`.

Optional duplicate summary in body:

- Human-readable move summary only.

Do not rely on body parsing for canonical state.

### 4.3 Graph API usage baseline

- Authentication: MSAL public client auth.
- Scopes: `Mail.ReadWrite`, `Mail.Send`, `User.Read`.
- Sync patterns:
  - Initial full pull for specific game folder.
  - Incremental polling with delta query or timestamp/query window.

---

## 5) Canonical Payload Schema

### 5.1 Envelope fields (required)

- `schemaVersion` (integer)
- `gameId` (GUID string)
- `messageId` (GUID string created by client)
- `parentMessageId` (GUID string or null for initial state)
- `ply` (integer, zero-based or one-based, choose once)
- `createdUtc` (ISO-8601)
- `senderAadObjectId` (string)
- `whitePlayerId` (string)
- `blackPlayerId` (string)
- `sideToMove` (`White` or `Black`)
- `status` (`Active`, `Checkmate`, `Resigned`, `Drawn`, `Abandoned`)
- `fen` (string)
- `moves` (array, cumulative history)
- `stateHash` (string, SHA-256 over canonicalized state)
- `prevStateHash` (string or null)

### 5.2 Move item fields (required)

- `ply`
- `san`
- `uci`
- `from`
- `to`
- `promotion` (nullable)
- `playedBy`
- `createdUtc`
- `resultingFen`

### 5.3 Optional fields

- `comment`
- `clock` metadata for future timed modes
- `variant` metadata if non-standard chess variants are added later

### 5.4 Example payload

```json
{
  "schemaVersion": 1,
  "gameId": "cc32780d-82da-4dd2-b0a4-cce87c74ce9a",
  "messageId": "5ab7f6ad-cb4a-49ec-9ed2-a10f50d6d595",
  "parentMessageId": "2a7dfb88-3f4d-4d84-a0af-59a3031cb918",
  "ply": 17,
  "createdUtc": "2026-07-31T21:30:05Z",
  "senderAadObjectId": "a2e2d75d-4bb2-4d8b-a677-c0ea1ff6d228",
  "whitePlayerId": "player-white",
  "blackPlayerId": "player-black",
  "sideToMove": "Black",
  "status": "Active",
  "fen": "r1bqkbnr/pppp1ppp/2n5/4p3/4P3/2N2N2/PPPP1PPP/R1BQKB1R b KQkq - 2 3",
  "moves": [
    {
      "ply": 16,
      "san": "Nc3",
      "uci": "b1c3",
      "from": "b1",
      "to": "c3",
      "promotion": null,
      "playedBy": "player-white",
      "createdUtc": "2026-07-31T21:30:05Z",
      "resultingFen": "r1bqkbnr/pppp1ppp/2n5/4p3/4P3/2N2N2/PPPP1PPP/R1BQKB1R b KQkq - 2 3"
    }
  ],
  "stateHash": "sha256:...",
  "prevStateHash": "sha256:..."
}
```

---

## 6) Chain Validation and Authority Rules

### 6.1 Message acceptance checks

A message is valid only if:

1. `schemaVersion` is supported.
2. `gameId` matches folder/game context.
3. `parentMessageId` exists (unless initial message).
4. `ply == parent.ply + 1`.
5. Move at `ply` is legal from parent position according to domain engine.
6. `fen` equals engine-calculated resulting position.
7. `stateHash` matches canonicalized payload.
8. Sender is the expected side to move for that `ply`.

### 6.2 Selecting current authoritative head

When more than one leaf exists:

1. Prefer leaf on longest valid chain.
2. If equal length, prefer chain with earliest first-arrival timestamp on local device.
3. Mark other leaves as branch conflicts.

This ensures deterministic local behavior while still surfacing branch conditions.

### 6.3 Conflict policy

Default policy for phase 1:

- Do not auto-merge branches.
- Show conflict UI with options:
  - Continue from selected head.
  - Fork as new game.
  - Archive conflicting branch.

---

## 7) Replay and Historical Message Behavior

### 7.1 Opening an old message

Opening any message in game folder should:

1. Parse payload and validate integrity.
2. Render board at that message `fen`.
3. Show move list up to that message `ply`.
4. Enable previous/next stepping within the message history.

### 7.2 Returning to live state

- Provide explicit `Return to Current` action.
- Current state is resolved from chain head selection rules.

### 7.3 Playing from historical state

Policy:

- If not current head, new move creates a branch and requires explicit confirmation.
- Optionally offer `Create Forked Game` to avoid accidental branch clutter.

---

## 8) Offline and Sync Model

### 8.1 Offline drafting

- Player may compose move offline from known current head.
- On reconnect, client revalidates against latest chain before send.

### 8.2 Stale-move handling

If local parent is no longer head:

- Block send by default.
- Offer user choice:
  - Rebase by loading latest head.
  - Send as branch.

### 8.3 Sync cadence

- Foreground polling every 15-30 seconds when game open.
- Immediate refresh when app regains focus.
- Manual `Sync Now` action.

---

## 9) Security, Trust, and Privacy

### 9.1 Identity

- Phase 1 supports both Entra tenant accounts and personal Microsoft accounts (MSA) through MSAL.
- Map participants by stable object IDs where possible, with fallback identifiers for account-type differences.

### 9.2 Integrity

- Validate hash chain (`prevStateHash` -> `stateHash`) at load time.
- Reject payloads failing cryptographic checks.

### 9.3 Data protection

- Store minimal local cache.
- Encrypt local cache at rest using platform facilities.

### 9.4 Permissions

- Request least-privilege Graph scopes needed for mailbox/game folder operations.

---

## 10) Cross-Platform Evolution Path

### 10.1 Shared layers

- `Chess.Domain`: board model, legality, check/checkmate, notation.
- `Chess.Sync`: Graph payload, validation, chain resolution.
- `Chess.Application`: use cases and orchestration.

### 10.2 Client layers

- Windows client first with native UX.
- Android/iPhone later using same domain/sync packages.

### 10.3 UI assumptions for portability

- Keep rendering/state-management UI-specific.
- Keep rules and chain logic UI-agnostic.

---

## 11) Implementation Plan

### Phase 0: Decision freeze and schema

- Finalize payload schema v1.
- Finalize chain-head and conflict policy.
- Define folder strategy in Graph.

### Phase 1: Domain extraction

- Port chess logic from legacy modules to pure .NET domain library.
- Add deterministic tests for legal moves, special moves, checkmate, replay.

### Phase 2: Graph sync core

- Implement auth, folder discovery/provisioning, send/receive payload.
- Implement chain validation and head selection.

### Phase 3: Windows app MVP

- Board UI, move list, replay controls.
- New game, make move, send move, receive move.
- Conflict surface and `Return to Current` experience.

### Phase 4: Hardening

- Corruption handling and recovery UX.
- Offline stale-move flows.
- Payload-size and performance tuning.

### Phase 5: Mobile clients

- Android and iPhone clients consume same sync/domain contracts.

---

## 12) Risks and Mitigations

1. Branch explosion from concurrent moves

- Mitigation: explicit stale detection and default block before send.

2. Message growth over long games

- Mitigation: optional compaction mode (checkpoint + tail) in schema v2.

3. Graph API throttling or transient failures

- Mitigation: retry with exponential backoff and sync checkpointing.

4. User confusion about historical vs live state

- Mitigation: clear UI badges for `Historical View` and `Current Head`.

5. Identity edge cases across account types

- Mitigation: normalize participant IDs and include fallback identifiers.

---

## 13) Decision Checklist Before Build

- Confirm Windows framework: WinUI 3 or MAUI desktop-first.
- Confirm game folder layout strategy in mailbox.
- Confirm payload storage format (JSON attachment preferred).
- Confirm tie-break rule for equal-length competing branches.
- Confirm default action when move is stale: block/rebase/branch.
- Confirm whether non-current historical moves create branch by default.
- Confirm and test cross-account play matrix in phase 1 (Entra-Entra, MSA-MSA, Entra-MSA).

---

## 14) Definition of Done for Roadmap B MVP

MVP is complete when:

1. Two Windows users can play a full game with legal move enforcement.
2. Each move produces a Graph message with full self-contained state payload.
3. Opening any old game message reconstructs board at that point.
4. Replay controls work for previous/next/jump and return-to-current.
5. Concurrent-move conflict is detected and clearly surfaced.
6. Chain integrity validation prevents invalid/tampered state acceptance.
