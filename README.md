# ERP 3-Way Match & Payment Automation

A tool I built as an inventory planner to automate purchase-order reconciliation — turning a manual, error-prone weekly task into an automated process that surfaces exceptions, prevents duplicate payments, and feeds Finance a clean line-level report.

**Result: weekly reconciliation fell from ~2 hours to ~30 minutes, and Finance adopted the output to clear split-invoice cases faster.**

---

## The Problem

Reconciling purchase orders against goods-received records and invoices (a "3-way match") was a manual, line-by-line task. Two things made it painful:

- **Upstream:** the planning logic generated multiple irregular shipments per vendor each week, inflating PO-line volume and creating date mismatches between goods-received and delivery records — which made downstream matching messy.
- **Downstream:** a single PO line was sometimes settled across multiple invoices ("split invoices"), which broke Finance's matching and slowed payment processing.

## What I Did

I fixed the root cause before automating anything:

1. **Redesigned the planning model** so each vendor's shipments consolidated onto a specific weekday. This cut total PO-line volume, narrowed the goods-received-to-delivery date gap, and made line matching far cleaner.
2. **Aligned records at the source** by getting vendors to share their own records at the point of each new order, so internal and vendor data agreed at the PO/goods-received level.
3. **Built the automation** (this repo) to handle what remained: import the ERP export, cleanse it, run the 3-way match, flag exceptions, age unpaid receipts, and catch duplicates.

## What the Tool Does

- **Imports and cleanses** the raw ERP export — strips cancelled, blank, and non-relevant lines; converts text-formatted numbers; looks up vendor names from codes.
- **Reshapes and sorts** the data into a clean working layout, deriving unit prices and received amounts.
- **Flags exceptions automatically** with conditional formatting: lines not yet received, goods-received vs. ordered-amount mismatches, and any line carrying a remark.
- **Ages unpaid receipts** — surfaces received-but-unpaid lines past 45 or 70 days, excluding intercompany/non-payable vendors.
- **Catches duplicate and split invoices** by building a composite key (vendor · PO · item · amount · invoice no.) and isolating any line that appears more than once.
- **Rolls up by invoice** so a single invoice spanning multiple PO lines is summed and reported clearly — the report Finance adopted to clear split-invoice cases.
- **Generates print-ready reports and a payment-request form.**

## Why It's Here

I'm not a developer — I'm a supply-chain planner who taught myself enough VBA to fix a process that was costing time and causing payment errors. I'm sharing it because it shows how I approach problems: find the root cause first, align the data, then automate the repeatable part. The same logic — defined rules, high record volume, zero tolerance for error — applies directly to adjacent processes like commissions calculation and any rules-based reconciliation.

## Technical Notes

- Written in **VBA (Excel)**. The matching, ageing, and duplicate-detection logic is in `ThreeWayMatch.bas`.
- Repeated operations are factored into shared routines (e.g. one filter-and-delete helper; the 45/70-day ageing routines are thin wrappers over a single shared function) to keep the logic in one place.
- Performance handled by toggling screen-updating and calculation mode around bulk operations.
- The source file path is read from a config cell, not hard-coded.

## A Note on Data

All identifying data has been removed. Vendor codes are placeholders, file paths come from a config cell, and there are no real company names, prices, or records anywhere in this repository. It demonstrates the *method*, which is standard procurement practice.
