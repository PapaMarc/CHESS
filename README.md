# PapaMarc/CHESS

_A historical artifact, a modernization workshop, and a place to reminisce about how far the internet has come._

## 🌱 Origins: The Exchange Server Sample Applications and the Application Farm

In the mid‑1990s, when Microsoft Exchange Server was still new, and the internet was still more potential than infrastructure, we created something called the **Exchange Application Farm** — a place intended to be _“fertile ground in which to grow bumper crops of applications.”_

The original AppFarm page (now preserved on the Wayback Machine) captured the spirit of that era:

**🔗 AppFarm (Wayback Machine)**  
https://web.archive.org/web/20000919125743/http://www.microsoft.com/TechNet/exchange/tools/AppFarm/appfarm1.asp

The idea was simple:

- Demonstrate **what** developers could build using Exchange Server as a platform
- Show **how**
  - store‑and‑forward messaging as a transport with private folders, and/or replicated, distributed public folders, with:
  - custom views - custom forms - custom properties

could serve as highly maliable respositories, for shared state, for rapidly developed and deployed applications.

So we seeded the ecosystem with real, working sample applications with modifiable shared source code.

Exchange Server 4.0 and 5.0 shipped with a set of Sample Applications (codename: **Butthead** — yes, _Butthead_, the code-generating WYSIWYG forms designer we built and shipped simultaneously was aptly codenamed: **Beavis**). And yes — as Seinfeld presenting “Butthead” to Bill Gates in the mid‑90s during peak _Seinfeld_ era, the moment practically wrote its own punchline. Bill gave me a look that landed somewhere between _“Should I fire this guy?”_ and _“This is actually kind of funny.”_ After a smirk, a head tilt, a long pause, and a head shake, he let me continue — and in the end, the samples shipped essentially intact.

These applications lived in the `\samples` directory of the server installation and were translated into more than 20 languages, helping developers worldwide by serving as both architectural and inspirational patterns — examples meant to spark imagination and encourage experimentation.

## ♟️ Hiring Aaron Bell — and the manifestation of the Chess sampleApp impl

One of those sample applications **Chess**, was coded by **Aaron Bell**. And the story of Aaron's touch and go at Microsoft is part of the charm.

As his hiring manager, i encountered Aaron when he was but a kid-- young enough that i had extra work to do with with Microsoft HR to ensure we weren’t violating child labor laws in putting him on the payroll.

He was ahead of his time, many a time... after Microsoft Aaron went off to college and laid down some lore and another application template which another Stanford-ite capitalized on in a big way several years later:
**🔗 The Stanford Daily: The Forgotten Social Network (2011)**
https://stanforddaily.com/2011/01/11/the-forgotten-social-network/

Years later, long after his time at Microsoft and Stanford, Aaron was profiled in _The New York Times_ as founder and CEO of AdRoll-- a wildly successful AI‑driven advertising platform used by 110K+ brands to deliver high‑performance, cross‑channel, online marketing. In his interview, he recounts his recollection of how his rapid fire stream of creative alternatives to my 'how do you know the light doesn't remain on after you close the fridge door' question, far outpaced numerous other candidates twice his age, and sealed 1/2dozen+ unanimous 'HIRE' recommendations from all on his day long interview loop.

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

In the early days of Exchange, Chess let us demonstrate how messaging could carry evolving state between participants. Today, it remains a beautiful example of how ideas from the past can be used to inform us of and carry us into the future on modern architectures.

## 🚀 PapaMarc/CHESS Today

This repository serves two purposes:

### **1. Archive**

It preserves the original Chess sample application exactly as it shipped with Exchange Server 4.0, capturing a moment when Microsoft was first weaving enterprise messaging into the fabric of the early Internet.

### **2. Workshop**

It plays with alternative ideas how that foundation can be modernized in current day:

- A web‑based version using modern transports
- A cloud‑distributed version using queues or durable messaging
- A real‑time version using WebSockets or SignalR
- A “retro mode” that still plays via email asynchronously, but using modern MIME, JSON payloads, and Microsoft's Graph APIs
- A version that uses replicated state in modern distributed systems (Redis, CRDTs, etc.)

Thinking through these possibilities is a way to reminisce — to revisit the foundations we built, the constraints we worked within, and the creativity that emerged from those constraints.

It’s also a way to stay connected to old friends, old ideas, and the spirit of the AppFarm... bridging into current day.

## 🧩 A work in progress

The AppFarm once described its goal as **cross‑fertilization of ideas** so that others could create more.  
GitHub repo PapaMarc/CHESS continues that tradition.

Chess was never just a game.

It was a message.  
A message about posibiilty, creativity, and connectivity.
A message carried forward, one move at a time, in its time.

AND a medium.
