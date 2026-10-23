# POV by Unrest — deployment update

## Files to update

1. Replace the repository's existing `index.html` with the supplied `index.html`.
2. Open the Google Apps Script project connected to the spreadsheet, replace the script source with `Code.gs`, save it, then deploy a **new version** of the existing Web app deployment.

The frontend uses this Apps Script Web App URL:

`https://script.google.com/macros/s/AKfycbxBaiSS-a-d1l8E6crR3w3hwn2EVJJoiE5DKuCR9vkZus9YNY1m9guyKT2J2a3_P7iR/exec`

Do not create a new Apps Script deployment URL unless you also update the URL in `index.html`.

## Apps Script deployment settings

- **Execute as:** Me
- **Who has access:** Anyone
- After saving `Code.gs`, use **Deploy → Manage deployments → Edit → New version → Deploy**.
- If the project is not yet connected to a spreadsheet, create/open the intended spreadsheet and ensure the script is bound to it before deploying.

## Verify after deployment

- Open the deployed site and load the opinions feed.
- Submit an opinion of 20–280 characters while the topic is open.
- React to an opinion and refresh the page to confirm the reaction persists.
- Choose a gender before submitting an opinion or voting in an open poll. Both selections are saved with the response. Poll results are hidden until the backend has at least 30 votes in the requested result slice.
- Add a campus food outlet and rate an eligible outlet. Food ratings and votes are shared through the Apps Script spreadsheet.

The Apps Script creates its data sheets automatically. The `topics` and `polls` sheets are configuration sheets; edit their rows to manage topic and poll configuration.
