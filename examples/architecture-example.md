---
title: Usage Service Reconciliation Improvement
subtitle: Architecture Point of View
author: Piyush
date: 29 Sep 2026
title_prefix: Systems & Infrastructure
footer: HP Confidential
---

## Problem Context

Usage Service publishes every transaction to CM and BIRD independently via SNS, so each transaction lands in one of three states.

| State | What happens | Today's DQ check |
|---|---|---|
| Healthy path | Delivered to CM and BIRD | Nothing to fix |
| Partially caught | In BIRD, missing at CM | Caught; fixed by email + manual resend |
| Blind spot | Never reached BIRD | **Never detected** |

> **Why this matters:** today's DQ check only compares data already inside BIRD's database, so it is structurally blind to transactions BIRD never received.

## Architecture Strategy

From detect-and-email to a self-healing loop.

1. **Detect**: daily recon lists every transaction missing in BIRD
2. **Sync**: generateXML backfills BIRD automatically
3. **Verify**: one DQ check after sync reports CM-only gaps
4. **Recover**: S3-triggered job resends to CM

> **Strategic point of view:** missing billing transactions are found and fixed automatically, with people involved only for exceptions.

## Benefits

### HP
- Protects revenue by closing the blind spot
- Cuts manual reconciliation effort across three teams
- Recovery becomes an automated daily cycle (by T+3)

### Customer / Partner
- More accurate, consistent invoices
- Faster, predictable resolution of missing usage
- Greater confidence in MPS billing

### Employee
- No more repetitive investigation and manual resends
- Lower on-call load from recurring incidents
- Time back for higher-value work

## Current vs. Proposed Flow

### Current (as-is)
1. **T+1 · US daily recon**: S3 dump of every transaction sent in the last 24h
2. **T+2 · BIRD daily recon**: compares SNS-received data vs. US dump; result stays internal
3. **T+2 · DQ check, in parallel**: blind to anything BIRD never received
4. **Email alert**: mixed-cause list emailed to the team
5. **Manual reprocess**: someone resends each transaction by hand
6. -

### Proposed (to-be)
1. [Unchanged] **T+1 · US daily recon**: same job; source of truth
2. [Enhanced] **T+2 · BIRD daily recon**: publishes an explicit "missing in BIRD" list
3. [New step] **T+2 · Auto-sync BIRD**: generateXML backfills, then signals completion
4. [Re-sequenced] **T+3 · Consolidated DQ check**: reports "not sent to CM" only
5. [Replaced] **Missed-transaction report → S3**: versioned, machine-readable
6. [New] **S3-triggered auto-reprocess**: idempotent resend to CM

## Target Architecture

```flow
Usage Service -> BIRD : 1. SNS publish
Usage Service -> Convergent Mediation [it] : 1. SNS publish
Usage Service --> BIRD : 3. generateXML sync
BIRD -> DQ Check : 4. after sync
DQ Check --> Missed-Txn Report [new] : 5. CM-only gaps
Missed-Txn Report [new] --> Auto-Reprocess Job [new] : 6. S3 PUT
```

<!-- notes: Numbers match the actions on the next slide. -->

## Outcomes

- **0** transactions silently missed after sync
- **T+3** worst-case recovery, down from open-ended manual follow-up
- **1** report condition: "not sent to CM"
- **0** manual resend steps in the normal path

## Technical Dependencies

| Dependency | Why it matters | How to address it |
|---|---|---|
| DQ check needs sync to finish first | In-flight transactions look missing purely from timing | Start the DQ check on a sync-complete signal, not a schedule |
| generateXML batch and throttle limits | Large batches load the BIRD DB; this now runs daily | BIRD team defines batch size, throttle and retries |
| Reliable US write-back | A failed write-back can cause a duplicate at CM | CM confirmation check before resend |
| S3 report schema and idempotency | A re-triggered event could resend twice | Versioned schema; idempotent on transaction key |
| No manual safety net | Job failures could go unnoticed | Fallback alert on job failure |
