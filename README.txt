SBTS FINAL

Patch43.2 (Cleanup1):
- Deduplicated repeated function declarations (Supabase helpers, Notifications inbox helpers, overlay helper).
- Cached Supabase client instance to avoid recreating client on each call.
- Fixed malformed duplicate function header for openBlindDetailsFromNotif.
- No feature removals; behavior should remain the same for local/offline mode.


PATCH 43.5 (Cleanup3)
- Added SBTS_UTILS helpers (qs/qsa/byId + safe JSON + LocalStorage wrapper)
- Replaced direct localStorage.getItem/setItem calls with SBTS_UTILS.LS.get/set (no behavior change)
- Goal: cleaner, safer maintenance; reduce 'blank index' risk from storage errors.


=== Patch 43.7 (Cleanup5) ===
- Boot refactor (no behavior change): sbtsBoot split into small init functions (state/config, theme, sidebar notifications link, routing, breadcrumbs, responsive handlers).
- Goal: easier maintenance + safer future upgrades.


PATCH 43.8 – Local Backup & Restore Improvements
- Create backup now: downloads JSON + saves a local snapshot.
- Added 'Save locally' button.
- Weekly auto-backup set to every 7 days, keeps last 10 snapshots.


Patch 43.9 – Backup Manager Upgrade
- Add label + notes when creating backups (file + local)
- Local backups list shows label/notes and adds Download + Delete actions
- Update backup info modal (weekly backups keep last 10)


---
Patch 43.10
- Added AUTO pre-action local backup before high-impact admin actions (Restore, Reset, Delete Area/Project, Permanent approval delete).
- Auto backups are stored locally in Backup Manager (last 10).


Patch 45.1.2:
- Refined Slip KPI cards to balanced KPI pills (no tall/stretchy cards)
- Same KPI pill style applied on Slip Blind page (Total/Completed/In Progress)


Patch 45.2.2:
- Recent Activity now shows friendly Blind references (Tag/Line/Short ID) instead of full UUID.


Patch 47.22: Rich Notification Details (Timeline + Activity + Context), adds lightweight activity logging and timestamps for read/archive/restore.


=== Patch 47.25 (Inbox UX) ===
- Gmail-style Toolbar in Details + Modal (Archive/Restore, Read/Unread, Done, Delete, Back/Close)
- Lazy Details: no auto-preview; preview opens only when clicking notification title.
- Removed broken openNotificationInInbox handler; title click now drives preview + second click opens modal.


Patch 47.31:
- Demo multi-user switcher (Acting as) uses Roles & Specialties Manager roles catalog.
- Smart empty-state hints for Views (My Actions / Current Project).
