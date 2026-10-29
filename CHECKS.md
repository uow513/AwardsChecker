# Awards Checker — what it checks

Reference for the validations in `AwardsChecker.bas`. Run the macro `RunAwardsCheck`
with the Awards file as the active workbook.

The macro looks for a worksheet named `Awards`, and falls back to the first sheet whose
name starts with `Awards` (a CSV opened in Excel is named after the file, e.g. `Awards (2)`).

## Columns it reads

| Column | Header | Used for |
|---|---|---|
| B | `AWC_DESCRIPTION` | Award — part of the NT baseline key |
| D | `AJB_DESCRIPTION` | First half of the job level |
| I | `ALV_DESCRIPTION` | Second half of the job level |
| K | `ADT_PAYITEMCODE` | Pay item code the rules are selected by |
| R | `PAY_EMPLOYEERATE` | Employee rate |
| S | `PAY_CLIENTRATE` | Client rate |
| T | `PAY_ADDITIONS` | Cleansed in step 1 |

Headers on R, S and T are verified at startup. If they don't match, the macro asks before
continuing. A row is skipped entirely when D, I and K are all empty.

## How it runs

Everything below happens in a **single pass** over the data rows. The pass cleanses column T,
builds the job level list, records the NT rate for each job level and award, and runs every
check that needs nothing but the row in front of it.

The one exception is the comparison against the NT rate: a row's NT row can sit below it in
the file, so that comparison can't be settled the moment the row is read. Those comparisons
are queued during the pass and resolved immediately after it, against the finished set of NT
rates. The queue is much smaller than the data, and the result is identical to checking each
row in place.

Because those queued results arrive last, the Validation Errors sheet is sorted by Awards row
number before it's handed back.

## Cleanse

Column T holds numbers that belong in column R.

| Situation | Action |
|---|---|
| T has a number, R is blank | Value moved to R, T cleared, R shaded yellow |
| T has a number, R already has a value | **Nothing moved** — flagged for manual review |
| T is blank or non-numeric | Left alone |

Every row with a number in T is listed on the **Cleanse Log** sheet, moved or not.

The validations read the cleansed value, so a rate moved out of T is checked like any other —
the move happens earlier in the same pass over the row.

## The job level

The job level is column D concatenated with column I, with no separator — e.g.
`A Bricklaying C3 NY12 38+0` + `S1`. The unique list appears on the **Job Levels** sheet.

### The NT baseline

Most rules compare a rate against the NT (Normal Time) employee rate. That baseline is
resolved per **job level *and* award**, not per job level alone, because the same job level
appears under more than one award with different rates — `A Carpentry C3 YR12 36+4 / S1`
exists under both BESIX and CFMEU, for instance. Each row is therefore compared against the
NT rate from its own award.

If a job level has no NT row at all, every rule that needs the baseline reports
"no NT row found for this job level" rather than passing silently.

## Validations

`<n>` below means digits, optionally with one decimal point — so `RDO38+2`, `RDO36+4` and
`RDO38+2.5` all match, while `RDOCO` does not.

| Pay item | Client rate | Employee rate |
|---|---|---|
| `AL` | must be 0 | = NT |
| `ALCASH` | must be 0 | = NT |
| `AL5` | must be 0 | = NT |
| `PL` | must be 0 | = NT |
| `WELLB` | not checked | = NT |
| `COMLEA` | not checked | = NT |
| `PHNW` | must be 0 | = NT × 2.5 |
| `RDO<n>+<n>` | must be 0 | = NT |
| `RDOACC<n>+<n>` | must be 0 | = −NT |

Applied to every row regardless of pay item:

| Check | Rule |
|---|---|
| Client vs employee | When the client rate is not 0, it must be **greater than** the employee rate. Equal rates fail. |
| Duplicate baseline | More than one `NT` row for the same job level and award is reported — the first one found is used as the baseline. |

### Comparison details

- **Tolerance.** Two rates within 0.0051 count as equal, so rounding to the cent passes:
  2.5 × 16.91 = 42.275 stored as `42.28` is accepted. Change `TOL` at the top of the module
  to tighten or loosen this.
- **Blank rates count as zero.** A blank employee or client rate is treated as 0, not skipped.
- **Zero-client rules are exact.** "Must be 0" is tested exactly, with no tolerance.

## Output

Three sheets, rebuilt from scratch on each run:

| Sheet | Contents |
|---|---|
| **Cleanse Log** | Every row with a number in column T: row number, job level, pay code, the value, the resulting employee rate, and whether it was moved |
| **Job Levels** | The unique job level list: both source columns, the award(s) it appears under, how many awards share it, pay item count, NT employee rate, error count, and OK/Errors |
| **Validation Errors** | One row per failed check, sorted by Awards row number: the row number, job level, award, pay code, the issue, the actual employee and client rates, the NT rate, and the expected value |

On the Awards sheet itself, failing rate cells are shaded red and cells moved in step 1 are
shaded yellow. Both shadings are cleared at the start of the next run.

## Caveats

- **Not yet run against the source export.** The logic was written from an inspection of the
  data, not verified by execution.
- **`PHNW` will likely produce many errors.** In the sampled rows its employee rate equals the
  NT rate rather than 2.5 × NT. Whether the data or the rule is wrong is a question for the
  award, not the tool.
- **The workbook is modified in place.** Step 1 writes to column R and clears column T. Keep a
  copy of the original export.
- **`RDOCO` (RDO Cash Out) is not covered** by the RDO rules — it has no `<n>+<n>` suffix. Only
  the general client-vs-employee check applies to it.
