# UI Design Specification — Healthsure

This version uses `Data Validation Tool v12 — Final Source Sum Insured` as the
functional base. Processing, validation, correction, relationship logic,
source-driven Sum Insured logic, and workbook generation are unchanged.

## Approved palette

- Primary: `#11a387`
- Secondary: `#0c725f`
- Dark: `#095244`
- Light background/highlight: `#e7e6f3`

## Visual treatment

- Gradient title/header bands use only the approved palette.
- Primary action buttons use a dark-to-primary gradient.
- Progress uses a dark → secondary → primary gradient.
- Cards use light-background surfaces with approved green borders.
- Hover states use palette gradients and subtle movement.
- Shadows use transparent shades derived from `#095244` / `#11a387`.
- No unrelated brand colors are intentionally introduced.
- Visible emoji icons were removed from the title/action controls to avoid
  introducing uncontrolled colors.

## User flow

Upload → Sheet selection → Business rules → Proceed → Processing progress →
Download final workbook.

The functional processing logic is not changed by this UI update.
