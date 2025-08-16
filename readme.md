# Gavel Strike
> A Google Sheets-integrated tool to simplify the Presiding Officer's role in Congressional Debate.

This tool is very much in beta. It may make mistakes. Please check its work.

An average Congressional Debate session has twenty competitors, who give around three speeches, totaling over 60 speeches that need to be tracked. On top of that, the format that Congressional Debaters use to keep track of speeches creates inefficiencies, namely the size of columns tracking competitors' names. And while the upper echelon of POs can overcome these, many beginner POs struggle in this regard.

This tool creates a standardized "calling order" that collapses precedency and recency into one easily-readable column to maximize PO efficiency at all levels. 

<!-- TODO: Add GIFs -->

### Usage
**To Open Calling Order:** Go to *Gavel Strike* → *Calling Order*.

**To Call A Speaker**: Click on the name of the speaker in the list. After a few moments, the list will be reordered.

### Installation
#### Extension Store
Gavel Strike is not available on the Extension Store for Google Docs. Follow the steps for manual installation.

#### Manual installation
These steps must be repeated for every PO sheet, unless you copy from a master doc:
1. Go to *Extensions* → *Apps Script*. If not already done, navigate to the newly opened tab.
2. In the `Code.gs` file, paste the contents of [this file](https://raw.githubusercontent.com/Uncodeable864/gavel-strike/refs/heads/main/Code.gs). 
3. Create a new file called `call-order.html`, and paste the contents of [this file](https://raw.githubusercontent.com/Uncodeable864/gavel-strike/refs/heads/main/call-order.html).
4. For the first launch only: Navigate back to `Code.gs` and push run.
5. Navigate back to your sheet, and there should be a tab called *Gavel Strike*! At this point, you may close the tab called "Apps Script".
6. Gavel Strike is now installed for this sheet only. There is no need to reinstall when opening the sheet again, and any copies of the sheet will also have Gavel Strike.
