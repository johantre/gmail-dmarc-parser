# gmail-dmarc-parser

This is a google apps script that parses the DMARC report mail sent from Microsoft.\
In Google Workspace this still needs to be deployed.

This repo itself is hosting the code, and is foreseen with a workflow that pushes your changes to your Google Apps Script environment. 

## ⚠️ Dependencies ⚠️ 
### The actual code deployed at Google
- The code does what it does, it only needs to be deployed to google by means of clasp.

### Deploying to Google
- mpn. package manager for getting clasp installed in your workflow. 
- clasp. The official tool to push code to Google Apps Script environment.\
This tool is not very stable, but the latest version allows you to get your change to Google by means of a "clasp push --force"\
It requires a github secret with credentials for google to execute though.\
Acquiring credentials happens by "clasp login" & following browser screen to allow the right permissions.\
This generates an .clasprc.json, which is to end up in your github secrets to run your workflow properly.

## Future enhancements

- nothing planned
