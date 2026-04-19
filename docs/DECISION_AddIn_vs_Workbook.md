# Decision Record: Add-In vs Workbook Distribution

**Date:** April 2026
**Decision:** Stay with .xlsm workbook. Do not build an add-in at this time.
**Status:** Active. Revisit when trigger conditions are met.

## Context

The RDK is a prototyping tool that demonstrates how to build config-driven financial models in Excel/VBA. The primary audience is developers who will study the architecture and implement a production version in their chosen technology. We evaluated whether to distribute the kernel as an Excel Add-In (.xlam) versus the current self-contained workbook (.xlsm) approach.

## Decision

**Stay with .xlsm.** The workbook approach is correct for a prototyping tool.

## Rationale

1. **Inspectability.** Developers need to see the kernel code alongside the domain code. A self-contained .xlsm puts everything in one place. An add-in hides the kernel one layer deeper — the opposite of what a prototype should do.

2. **Zero install.** Setup.bat + .xlsm works immediately. No add-in installation, no admin privileges, no "enable this add-in in Excel settings" friction.

3. **Mac compatibility.** VBA add-ins have limitations on Mac. COM add-ins don't work on Mac at all. The .xlsm approach works on both platforms (with manual setup on Mac).

4. **Architecture readiness.** The kernel/domain separation work (zero domain leaks, config-driven everything, domain contract) means we CAN migrate to an add-in later without re-architecting. The prerequisite work is done.

5. **Premature optimization.** An add-in solves distribution and versioning problems we don't have yet. We have one model (insurance FM) used by one team. When we have 3+ models with version drift, the calculus changes.

## Pros of Add-In (for future reference)

| Pro | Value | When It Matters |
|-----|-------|-----------------|
| Single kernel, many models | Fix once, all models updated | 3+ active models |
| Cleaner workbooks | .xlsx instead of .xlsm, smaller files | External distribution |
| Professional distribution | Ribbon UI, no VBA security warnings on workbook | End-user product |
| Version control simplicity | One kernel version, no copy drift | Multi-developer team |
| IP protection | Kernel code hidden (password-protected .xlam or compiled COM) | Commercial distribution |

## Cons of Add-In

| Con | Impact | Mitigation |
|-----|--------|------------|
| Installation friction | Users must install add-in before opening any model | Installer package, but adds complexity |
| ThisWorkbook references break | ~200 occurrences of ThisWorkbook must change to workbook reference parameter | Mechanical refactor, ~2-3 days |
| State isolation | Module-level variables leak between workbooks if two models open | Workbook-keyed dictionaries instead of module-level vars |
| DomainOutputs handoff | Application.Run across add-in boundary has restrictions | Test thoroughly, may need alternative handoff |
| Two-file distribution | Always ship add-in + workbook/config | Package manager or installer |
| Debug complexity | Stepping through add-in code while model workbook is active | VBA editor handles this, but it's clunky |
| Mac limitations | VBA .xlam works but COM add-in does not | VBA-only add-in, no COM |
| Excel version sensitivity | Add-in behavior varies across Excel 2016/2019/365 | Test matrix across versions |
| Testing complexity | Can't just import .bas files and test | Need add-in loaded + test workbook open |

## Trigger Conditions to Revisit

Revisit this decision when ANY of these become true:

1. **3+ active models** with kernel copy maintenance becoming a burden
2. **External user distribution** where install experience matters
3. **IP protection** becomes a requirement (commercial distribution)
4. **Ribbon UI** is needed for professional presentation
5. **Multi-developer team** where kernel version drift causes bugs

## Migration Cost (Estimated)

When the time comes, the migration work is:

| Task | Effort | Risk |
|------|--------|------|
| Change ThisWorkbook to workbook parameter | 2-3 days | Medium (mechanical but high-volume) |
| Add state isolation | 1 day | Medium (cache/dictionary changes) |
| Build ribbon XML | 0.5 day | Low |
| Test DomainOutputs handoff across boundary | 0.5 day | High (BUG-034 workaround may not cross boundary) |
| Test across Excel versions (2016/2019/365/Mac) | 1-2 days | Medium |
| Update Setup.bat / installer | 0.5 day | Low |
| **Total** | **5-7 days** | |

## Alternative Considered: Both Add-In and .xlsm

Rejected. Maintaining two distribution formats doubles testing surface, creates confusion about which is canonical, and provides no benefit for the prototyping use case.

## Alternative Considered: .xlsb (Binary Workbook)

Not yet evaluated. Binary workbooks load faster than .xlsm and support VBA. Worth testing as a zero-effort performance improvement. No architectural change required — just save as .xlsb instead of .xlsm.

## Related Documents

- [MAC_COMPATIBILITY.md](MAC_COMPATIBILITY.md) — What works on Mac, what doesn't, workarounds
- [RDK_Architecture.md](RDK_Architecture.md) — Four-layer stack, domain contract, config system
