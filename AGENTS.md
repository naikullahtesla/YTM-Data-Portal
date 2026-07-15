# Carrefour YTM Data Portal

## Project Overview
React (Vite) single-page app for Carrefour's YTM (Year-to-Month) basket intake workflow. Parses monthly `ExportDetail` Excel files, enriches order lines from a Supabase product catalogue, flags issues (missing PCB, duplicates, below MOQ), and exports database-ready SOLID/PRINT sheets or pushes directly to Supabase.

## Tech Stack
- **Frontend:** React 18, Vite 5, JSX (no TypeScript)
- **Styling:** Inline CSS-in-JS (template literal `CSS` string), CSS variables for light/dark theme
- **Data:** Supabase (PostgreSQL), xlsx for Excel parsing
- **Icons:** lucide-react
- **Auth:** Plain-text passwords stored in `users` table, session in localStorage

## Commands
```bash
npm run dev        # Start dev server
npm run build      # Production build (dist/)
node seed-admin.mjs <email> <password> [name] [--force]  # Create/overwrite admin user
```

No linter or test suite is configured.

## Environment
`.env.local`:
```
VITE_SUPABASE_URL=https://xxx.supabase.co
VITE_SUPABASE_ANON_KEY=sb_publishable_xxx
```

## File Structure
```
src/
  App.jsx          # Everything: components, CSS, logic, helpers (~1600 lines)
  supabase.js      # Supabase client init (6 lines)
seed-admin.mjs     # CLI script to create admin users
migrate.sql        # Full DB schema: tables, triggers, indexes, RLS
Reference DB/      # CSV source data (Dyed.csv, Prints.csv)
```

## Database Schema (Supabase)

### Tables
| Table | PK | Purpose |
|---|---|---|
| `users` | `email` | Auth: email, password (plain), full_name, role (admin/user) |
| `colors` | `color_code` | Lookup: color_code ↔ color_1 name |
| `dyed_products` | `product_code` | Product catalogue for solids |
| `print_products` | `product_code` | Product catalogue for prints |
| `dyed_orders` | auto | Order lines for solids (FK → dyed_products) |
| `print_orders` | auto | Order lines for prints (FK → print_products) |

### Key Relationships
- `dyed_products.color_code` → `colors.color_code` (FK, auto-resolved by trigger on `color_1`)
- `print_products.color_code` → `colors.color_code` (FK, auto-resolved by trigger on `color`)
- `dyed_orders.product_code` → `dyed_products.product_code` (FK, triggers auto-fill product fields)
- `print_orders.product_code` → `print_products.product_code` (FK, triggers auto-fill product fields)
- Colors cascade: updating `colors.color_code` propagates to all product tables via trigger

### Triggers (auto-fill, don't set these manually)
- `trg_resolve_color_code` — resolves `color_code` from `colors` when `color_1` is set on dyed_products
- `trg_resolve_print_color_code` — same for print_products (keyed on `color`)
- `trg_propagate_color_code` — cascades color_code changes from colors to both product tables
- `trg_resolve_dyed_order_products` — fills all product fields on dyed_orders from dyed_products
- `trg_resolve_print_order_products` — fills all product fields on print_orders from print_products

## App Architecture (src/App.jsx)

### Core Data Flow
1. User uploads basket `.xlsx` → parsed by xlsx library
2. `buildLine()` enriches each row: season, entity, article, route (solid/print), dates, weights
3. Supabase data fetched on load → `buildDbFromSupabase()` builds code index + dup key set
4. Duplicates detected via `dupKey(code, color, basket)` against DB and within file
5. Export pushes to Supabase (products upserted, orders inserted) or downloads Excel

### Key Functions
- `buildLine(row, dbIndex)` — core enrichment per basket row
- `buildDbFromSupabase(data)` — builds `{ codes, keys }` from 4 Supabase tables
- `exportToSupabase(lines)` — upserts products, inserts orders, handles colors
- `solidRow(L)` / `printRow(L)` — maps enriched line to DB column format
- `ENTITY_MAP` — maps 4-char entity codes (FRCA, BECA, etc.) to customer names

### Components
- `LoginPage` — email/password auth, no signup
- `App` (default export) — main app: header, nav pills, basket upload, data table, KPIs, export
- `ColorsManager` — CRUD + export/upload CSV for colors table
- `ProductsManager` — CRUD + export/upload CSV for dyed/print products (shared, parameterized)
- `SettingsPage` — admin: user management (add/edit/delete/change password); non-admin: change own password

### Pages (nav pills)
- **Basket** — main workflow: upload, process, export
- **Colors** — manage color lookup table
- **Dyed Products** — manage dyed product catalogue
- **Print Products** — manage print product catalogue
- **Settings** (cog icon) — user management (admin) + password change

### State Management
All state lives in `App` component via `useState`. No Redux, no context. Key states:
- `lines` — processed basket lines
- `dbIndex` / `dbKeys` — Supabase data for enrichment + duplicate detection
- `sbStatus` — idle/loading/ready/error
- `page` — current nav page
- `user` — logged-in user from localStorage session

## Auth System
- **Passwords:** plain text stored in `users.password` column
- **Session:** `localStorage` key `bic_session` stores `{ email, name, role }`
- **Roles:** `admin` (full access + user management) and `user` (basket + data + own password only)
- **No signup flow** — users created by admin via Settings page or `seed-admin.mjs`

## Styling Conventions
- All CSS in a single `CSS` template literal inside `App.jsx`
- CSS variables for theming (`.bic` light, `.bic.dark` dark mode)
- No external CSS files, no CSS modules, no Tailwind
- Components use inline `style={}` for one-off overrides
- Borderless design: `border: none` on all cards, inputs, buttons
- Toast notifications: fixed bottom-right, auto-dismiss after 3s

## Important Notes
- `migrate.sql` must be run in Supabase SQL Editor before first use
- After schema changes, re-run `migrate.sql` and re-seed users
- The `salt` column was removed — if migrating from old schema, run the ALTER statements in the chat
- No build-time type checking — all JS, no TypeScript
- Bundle is ~850KB (xlsx library is the main contributor)
