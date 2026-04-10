# HAL Build Log

## 2026-03-30

### Project direction
- Confirmed HAL should remain local-first for now.
- Decided that a Teams personal tab is optional rather than required.
- Shifted Microsoft integration strategy toward Outlook calendar access instead of centering the product around a Teams app surface.
- Kept long-term direction open for a future desktop app wrapper to support stronger microphone behavior.

### Core product architecture already in place
- Built a local-first HAL experience centered around microphone-driven capture and a compact command panel.
- Added structured sections for `My Day`, `Tasks`, `Ideas`, `Meeting Notes`, and `Reminders`.
- Added local persistence for notes, tasks, ideas, meeting notes, reminders, quick links, theme choice, and learned HAL preferences.
- Added adaptive interpretation scaffolding with local correction memory and an optional OpenAI-backed interpretation path.

### UI and workflow changes completed
- Reworked layout around:
  - command capture at the top
  - `Meeting Notes` and `Ideas` on the left
  - `My Day` and `Tasks` on the right
- Added dark mode and sunrise-inspired light mode with icon toggles.
- Replaced text-heavy controls with icon-based controls where appropriate.
- Added the HAL red-eye microphone interaction and listening pulse states.
- Simplified the capture UI to focus on microphone input and a single save action.

### Tasks
- Added multiple task lists with `Regular Tasks` as the default list.
- Added popup task-list windows for viewing and managing tasks.
- Added task editing, deletion, reassignment to lists, optional due time, and completed-task handling.
- Added natural-language task parsing for due dates like `today`, `tomorrow`, `tonight`, weekdays, and other common phrases.
- Added duplicate protection for tasks.
- Added past-due task handling with a popup and quick reschedule controls.
- Added quick reschedule controls to `My Day` tasks.
- Changed completion behavior so only task checkboxes mark tasks complete.

### My Day
- Built `My Day` to show schedule items and tasks for the selected date.
- Added manual event creation.
- Added past-due task access from `My Day`.
- Added Outlook sync scaffolding so the selected day can eventually pull Microsoft calendar events once sign-in is configured.

### Ideas
- Added compact Ideas widget plus popup list and full editor dialog.
- Added archive and delete actions.
- Added rich-text editing tools including bold, italics, underline, lists, clear formatting, and compact font/size controls.
- Added export actions for Word, PowerPoint, and Excel-style outputs.

### Meeting Notes
- Added searchable meeting notes with compact previews.
- Added formatting logic for bullet-style prompts such as `make a note to talk about the following`.
- Removed the automatic `Follow-Up` section from generated meeting notes.

### Reminders
- Added one-time and recurring reminder support.
- Added recurring options for daily, weekly, monthly, weekdays, all-days, and ordinal weekday-of-month patterns.
- Added reminder editing and deletion.
- Added voice creation, deletion, and editing for reminders.
- Added follow-up questioning when reminder time is missing.
- Added Teams and SMS delivery scaffolding.
- Moved notification settings behind a gear icon into a dedicated reminder settings dialog.

### Voice and assistant behavior
- Added spoken HAL responses with a toggle.
- Tuned voice preference toward calmer male browser voices when available.
- Added end-of-prompt voice commands like `HAL, save it`.
- Added correction handling so recent items can be reclassified after mistakes.
- Added conversational correction context to better handle back-and-forth fixes.
- Tightened task and meeting-note intent routing so explicit task-list references and note phrasing stay out of Ideas more reliably.
- Added a real `meeting-update` interpretation path so prompts that add to an existing meeting note append to the matched note instead of creating a new idea.
- Improved correction handling for wrong-task-list cases so `move it to <list>` can reassign the existing task instead of creating a duplicate.
- Reduced stale correction bleed-through by clearing correction context on new prompt handling and preferring the most recent mistaken item during corrections.
- Added portable local learning memory for task-list corrections and meeting-note target corrections so HAL can reuse your fixes later without new code changes.
- Applied learned routing preferences before AI fallback so repeated corrections can steer HAL into the right task list or existing meeting note automatically.

### Integrations and backend scaffolding
- Added Microsoft auth scaffolding with local backend support.
- Added Graph-backed Teams meeting creation scaffold.
- Added Teams activity-feed reminder notification scaffold.
- Added SMS scaffold with Twilio-ready backend configuration and test route.
- Added Outlook calendar day-sync backend route for `My Day`.

### Current blocker / note
- Outlook or Teams integrations still require Microsoft Entra app configuration in `.env`.
- If IT does not allow custom Entra app registration or consent, Microsoft Graph-based integrations may not be possible in this environment.

### Google calendar scaffold
- Added a `Schedule Settings` gear to the `My Day` card.
- Added a calendar source selector for `Local only`, `Outlook calendar`, or `Google calendar`.
- Added Google calendar display-mode settings for `Work`, `Personal`, or `Both`.
- Added saved labels for Google `Work` and `Personal` calendars to support later mapping.
- Added a backend Google calendar status scaffold so HAL can report whether Google OAuth credentials have been configured.

### CSV calendar import
- Added a local CSV calendar source for `My Day` that reads `calendar_export.csv` from the app workspace through a backend route.
- Added CSV day filtering and time sorting so imported schedule items render cleanly on the selected date.
- Added auto-refresh behavior for the CSV source on browser reload and when the `My Day` date changes.

### Backup and migration safety
- Added a `Backups` control in HAL so local data can be protected before any online migration work.
- Added downloadable JSON backups that capture tasks, ideas, meeting notes, reminders, quick links, calendar settings, and HAL learning memory.
- Added local snapshot creation through the backend, which writes timestamped backup files into a `backups` folder inside the HAL workspace.
- Added restore-from-backup support so a saved backup file can replace the current local HAL data if something goes wrong during future migration work.

### First online-migration architecture step
- Added a local server-state mirror at `data/hal-state.json`.
- HAL now keeps browser local storage as its immediate safety layer while also mirroring its state to the backend.
- On startup, HAL can recover from the server mirror if the browser copy is empty, which starts the move toward a backend-owned HAL state without risking the current local experience.
- Extracted that mirror into a reusable server-side state store module so the app is less tightly coupled to a single JSON file implementation.
- Added env-configurable storage roots plus a starter [render.yaml](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\render.yaml) so HAL has a concrete first hosted path when you're ready.

## Logging policy going forward
- Continue appending major architecture, workflow, and integration changes here as HAL evolves.
