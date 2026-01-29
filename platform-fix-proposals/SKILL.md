---
name: platform-fix-proposals
description: Creates elite consulting proposals for Platform Fix engagements. Triggers when creating proposal documents for new clients, including discovery call follow-ups, pricing options, and professional DOCX generation. Incorporates Alex Hormozi techniques (value stack, bonuses, scarcity, urgency, risk reversal) and Platform Fix branding.
---

# Platform Fix Proposal Creator

## Quick Start

1. Read `references/pricing-guidelines.md` for tier structure and rates
2. Read `references/proposal-structure.md` for section-by-section template
3. Read `references/hormozi-techniques.md` for persuasion framework
4. Copy `assets/proposal-template.js` as DOCX generator base

## Required Information

Gather from discovery call:

| Field | Example |
|-------|---------|
| Company name | Weisshorn C5I |
| Contact name | Benoit Perroud |
| Location | Switzerland |
| Current state | 3 teams, scaling to 30 |
| Pain points | Single point of failure, 20-40% productivity loss |
| Desired outcomes | Self-service onboarding, no heroics |
| Timeline | Q1 2026 start |
| Blockers | Classified environment access |

## Standard Tiers

| Tier | Price | What They Get |
|------|-------|---------------|
| Foundation | £35,000 | 6-week transformation |
| Momentum | £50,000 | Foundation + 6mo advisory + £5k bonus |
| Accelerator | £70,000 | Momentum + 4 days/month + £8k bonus |
| Partnership | £100,000 | Accelerator + 8 days/month + £15k bonus |

**Price gaps:** ~40-45% between tiers (logical, not jarring)

**Standalone Audit:** £7,500 (ad-hoc only, not advertised). 60-day upgrade credit applies.

**Custom pricing:** +£5-10k for international travel, classified environments, compliance, urgency.

## The Guarantee

> **SAVINGS GUARANTEE:** By Week 2, we identify £100k+ in annual savings — or we refund your entire deposit and you keep the audit.
>
> **DELIVERY GUARANTEE:** We complete in 6 weeks — or we continue at no cost until done.
>
> **You cannot lose money on this engagement.**

## "Ideal For" Framing

Frame around goals, not deficiencies:

| Tier | Say This | Not This |
|------|----------|----------|
| Foundation | "Teams ready to move fast and own the outcome" | "Teams with 3+ engineers" |
| Momentum | "Teams who want expert guidance as they scale" | "Teams with 1-2 engineers" |
| Accelerator | "Teams in active transformation who need hands-on support" | "Teams with 0-1 engineers" |
| Partnership | "Teams who need a Platform Director without the £200k salary" | "Enterprises needing leadership" |

## Bonuses (Not "Free Months")

| Tier | Bonus | Value |
|------|-------|-------|
| Momentum | Priority Onboarding Package | £5,000 |
| Accelerator | Migration Readiness Audit | £8,000 |
| Partnership | Executive Briefing Package | £15,000 |

**Never say "free month." Frame as included bonus.**

## Hormozi Checklist

- [ ] Dream outcome stated
- [ ] Value stack: total > price by 30%+
- [ ] Bonuses with £ values (not "free")
- [ ] Scarcity: "2 clients per quarter"
- [ ] Urgency: Bonus expiry + cost of delay
- [ ] Guarantee: Full deposit refund
- [ ] "Ideal for" = goals, not deficiencies
- [ ] Price anchoring: Show tiers, recommend one
- [ ] Reason why: Market entry, timeline fit

## Document Design

| Element | Value |
|---------|-------|
| Primary | Navy #0D1B2A |
| Accent | Gold #C9A227 |
| Success | Green #059669 |
| Warning | Orange #D97706 |
| Font | Calibri / Calibri Light |
| Page | A4 |

## Writing Rules

**Do:** Short sentences. Active voice. Specific numbers. British spelling.

**Don't:** Em dashes. Waffle. "Free month." Apologise for price.

## Output Checklist

- [ ] Client name correct
- [ ] Pricing matches tier (or custom justified)
- [ ] Payment terms sum to total
- [ ] Dates correct
- [ ] No em dashes
- [ ] Value stack total > price by 30%+
- [ ] Guarantee refund = deposit amount
- [ ] Bonuses framed correctly (not "free")
- [ ] "Ideal for" uses goal framing

## Outputs

| Type | Location |
|------|----------|
| Proposal | `/mnt/user-data/outputs/[Client]_Proposal.docx` |
| Email | In conversation |