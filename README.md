# PapaMarc/CHESS

_A historical artifact, a modernization workshop, and a place to reminisce about how far the internet has come._

## 🌱 Origins: The Exchange Application Farm

In the early‑1990s, when Microsoft Exchange Server was still new and the internet was still more potential than infrastructure, we created something called the **Exchange Application Farm** — a place intended to be _“fertile ground in which to grow bumper crops of applications.”_

The original AppFarm page (now preserved on the Wayback Machine) captured the spirit of that era:

**🔗 AppFarm (Wayback Machine)**  
https://web.archive.org/web/20000919125743/http://www.microsoft.com/TechNet/exchange/tools/AppFarm/appfarm1.asp

The idea was simple:

- Demonstrate what developers could build using Exchange’s replicated store
- Show how replicated, distributed (private and public) folders could act as shared state
- Highlight custom forms, custom properties, and store‑and‑forward messaging as a transport for applications
- Seed the ecosystem with real, working sample applications with modifiable source code

Exchange shipped with these samples not as toys, but as architectural patterns — examples meant to spark imagination and encourage experimentation.

## ♟️ Hiring Aaron Bell — and the Birth of the Chess Sample

One of those sample applications **Chess**, was written by **Aaron Bell**.

And the story of how Aaron arrived at Microsoft is part of the charm.

I hired Aaron when he was still a child — young enough that i had to work with Microsoft HR to ensure we weren’t violating child labor laws in putting him to work.

Years later, long after his time at Microsoft, Aaron was profiled in _The New York Times_ as founder and CEO of AdRoll. That article mentions the refrigerator‑light interview — but it happened _decades_ earlier, during our interview together and i still recall those he also recalls.

**🔗 NYTimes Profile of Aaron Bell (2016)**  
https://www.nytimes.com/2016/03/20/business/aaron-bell-of-adroll-the-truth-may-hurt-but-it-also-heals.html

## 📨 The Original Chess Application

The original Chess sample — preserved in this repo’s `Archive/` folder — is a Visual Basic implementation of a **custom‑form, email‑based, store‑and‑forward chess game**.

Two players could play asynchronously over email.  
Each message contained:

- The full move history up to that point
- The current board state
- The custom form used to render the board and submit the next move

It was clever, lightweight, and a perfect demonstration of what Exchange’s architecture made possible. It showed how a message could be more than communication — it could be a vessel for computation and shared state.

## 🔧 Why Chess?

Chess has always been fertile ground for experimentation:

- Structured and deterministic
- Expressive and endlessly deep
- Social — a way to stay connected with friends and colleagues
- A perfect testbed for distributed systems, synchronization, and state transfer

In the early days of Exchange, Chess let us demonstrate how messaging could carry evolving state between participants. Today, it remains a beautiful example of how ideas from the past can inspire modern architectures.

## 🚀 PapaMarc/CHESS Today

This repository serves two purposes:

### **1. Archive**

It preserves the original Chess sample application exactly as it shipped with the first version of Exchange Server — a historical artifact from the formative era of Microsoft’s groupware platform.

### **2. Workshop**

It explores how that idea could be modernized:

- A web‑based version using modern transports
- A cloud‑distributed version using queues or durable messaging
- A real‑time version using WebSockets or SignalR
- A “retro mode” that still plays via email, but using modern MIME, JSON payloads, or Graph APIs
- A version that uses replicated state in modern distributed systems (Redis, CRDTs, etc.)

Thinking through these possibilities is a way to reminisce — to revisit the foundations we built, the constraints we worked within, and the creativity that emerged from those constraints.

It’s also a way to stay connected to old friends, old ideas, and the spirit of the AppFarm.

## 🧩 A Living Memory

The AppFarm once described its goal as **cross‑fertilization of ideas** so that others could create more.  
PapaMarc/CHESS continues that tradition.

Chess was never just a game.  
It was a message.  
A message about what was possible.  
A message about what was coming.  
A message carried forward, one move at a time.
