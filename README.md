# GTO Operations Portal

Staff portal for GTO (Global Travel Organisation): log-in status (Work / Break / Log out), leave requests with two-step approval, schedules from the rota, and shift swaps.

## How it works
- `index.html` – the portal page. Staff sign in with their Microsoft work account. Add `?demo` to the address for a demo with sample people.
- `api/index.js` – the server (Vercel). It checks who is signed in against the **Access** list and their role, then reads and writes the Excel files **as the opsportal account**. Staff never get access to the files themselves.
- The Excel files are the database: `GTO Portal Data.xlsx` (opsportal OneDrive), the GTO Leaves Tracker and the rota.

## Settings (Vercel → Project → Settings → Environment Variables)
| Name | Value |
|---|---|
| `CLIENT_SECRET` | Client secret from Entra → App registrations → GTO Operations Portal → Certificates & secrets |
| `BLOB_READ_WRITE_TOKEN` | Added automatically when a Blob store is connected |

After deploying, open `/api/setup` once and sign in as opsportal@gto.org.uk to connect the portal to the files.
