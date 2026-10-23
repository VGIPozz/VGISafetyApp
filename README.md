# The Violin Group – Site Safety

A Microsoft Teams tab for crews to complete on-site safety documents. The crew lead fills in the job details once on the home page, ticks today's work activities, and the app lists the safety documents required for that job. Each form opens with the job details already filled in.

Plain HTML/CSS/JavaScript, with no build step. Hosted on GitHub Pages.

## Files

| File | What it is |
|---|---|
| `TheViolinGroupSafetyApp.html` | Home page: job details, crew, emergency info, work activities, safety document checklist |
| `index.html` | Redirects the site's main address to `TheViolinGroupSafetyApp.html` |
| `form.html` | "Coming soon" page shown for any form that isn't built yet (shows the job details it will receive) |
| `config.html` | Teams tab setup page (paste the Power Automate flow address) |
| `assets/docs.js` | **The list of work activities and safety documents.** Edit this to add, rename or remove forms, or change which activities make a form required |
| `assets/app.js` | Shared code: saves the job on the device, form status, flow address, Teams theme and sign-in name |
| `assets/app.css` | Shared styles: The Violin Group gold (#8f734b), black and white, in light and dark |
| `assets/logo-light-text.png` | Logo with white wordmark, used on the black page headers |
| `assets/logo-dark-text.png` | Logo with black wordmark, for white backgrounds (Teams setup page, future PDFs) |
| `forms/` | Fillable forms go here as they're built |

## Adding a fillable form (later)

1. Build the form as `forms/<id>.html`. Load `../assets/docs.js` and `../assets/app.js`, and read the job with `SafetyApp.Job.get()`. Fields: `site, address, customer, contractor, when, jobNo, scope, leadName, leadTitle, musterPoint, emergencyContact, activities`, plus `SafetyApp.Job.workers(job)` for everyone on site.
2. After a successful submit, call `SafetyApp.Job.markDone('<id>', { fileName })` so the home page shows it as done.
3. In `assets/docs.js`, add `page: 'forms/<id>.html'` to that document.

## Publish on GitHub Pages

1. In the public repository **VGIPozz/VGISafetyApp**, upload everything in this folder to the top level (not inside a folder).
2. Go to **Settings → Pages**, choose **Deploy from a branch**, pick `main` and `/ (root)`, then click **Save**.
3. After a few minutes the app is live at `https://vgipozz.github.io/VGISafetyApp/`.

## Add to Teams

- In the Teams app manifest, set the configurable tab's `configurationUrl` to `https://vgipozz.github.io/VGISafetyApp/config.html`, and add `vgipozz.github.io` to `validDomains`.
- When someone adds the tab to a channel, they paste the Power Automate flow address on the setup page. It's stored with the tab and is never written into this public code.

## How job details are kept

The job is saved on the device (browser storage inside Teams) so it's still there when moving between the home page and each form, or when the phone switches apps. It stays until someone taps **Start a new job**. If the saved job is from an earlier day, the home page asks whether to start a new one.
