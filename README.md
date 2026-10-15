# Lab operations dashboard in Excel

> **In short:** Lab managers needed one place to see how each of nine sites was doing every week: how many tests were run, how many results were on time, whether there were enough staff, and whether the machines were working. I built this Excel dashboard so they can pick a week, site or department and every number updates. It showed that staff shortages kept happening at the same sites in the same weeks, which changed how rosters were planned.

A weekly operations dashboard for a multi-site diagnostic lab, used by lab managers in live management reporting. It tracks four KPIs across nine sites: tests run, unique patient requisitions, TAT compliance against a 90% target, and staff adequacy, plus analyser availability.

**Tools:** Excel (SUMIFS, AVERAGEIFS, COUNTIFS, nested IF, conditional formatting). No macros.

Site names, analyser names and figures in this copy are placeholders ("Site A", "Sample Brand"); the structure and formulas are the ones in use.

## The problem

Getting a cross-site picture meant opening several spreadsheets, comparing numbers by hand and hoping nobody had broken a formula upstream. Managers spent their time assembling data instead of acting on it.

## How it works

Four sheets, each with one job:

| Sheet | Who touches it | What it does |
|---|---|---|
| `DASHBOARD` | Managers | KPI cards, tests and TAT by site and department, staffing gaps, analyser status |
| `DATA ENTRY` | Whoever does the weekly update | Three sections: weekly test volume and TAT, analyser register, staffing register |
| `SELECTIONS` | Managers | One place to set year, month, week, site and department; every figure on the dashboard follows it |
| `PIVOT DATA` | Nobody | All-time totals the dashboard compares against |

Every KPI card compares itself with the previous week automatically:

```excel
=IF(A57=0, "No prior week data",
 IF((A52-A57)>0, "up +"&TEXT(ABS(A52-A57),"0.0%")&" vs prev week",
 IF((A52-A57)<0, "dn -"&TEXT(ABS(A52-A57),"0.0%")&" vs prev week",
 "no change vs prev week")))
```

Analyser availability is calculated from logged downtime and repair hours, `(168 − downtime − repair) ÷ 168` per week, so an instrument that keeps going down shows up as a falling percentage rather than an anecdote.

## What it changed

With weekly cross-site data in one place, staffing gaps turned out not to be random: they clustered at particular sites in particular weeks. That made them something rosters could plan around.

## Files

`Lab_Operations_Dashboard_Portfolio.xlsx`: open the `DASHBOARD` sheet, then change the filters on `SELECTIONS`.

This workbook was the starting point for the [TAT analytics platform](https://github.com/Yusuf-Tiamiyu/lab-tat-platform-demo), which automated the TAT side of it end to end.
