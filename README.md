# Project Catalyst Review of Reviews Tools

This project is simple collection of scripts to speed-up the "Review of
Reviews" stage in Project Catalyst and integrate the set of principles
described in [THE YELLOW AND RED CARD THING](https://docs.google.com/document/d/1D04NRH1U_ZhAqaA3Gdk0WZqDcYkXoPZpC_15t5QyX14).
This is an attempt to facilitate and improve the decision-making process for
vCAs who are reviewing CA reviews in Fund 4.

In Fund 4, as for Fund 3, this process is based on Google Sheet and these
tools interact with it using `gspread`.

## Disclaimer
This is not the most elegant script collection in the world. It was just a quick
a dirt development to get the tool ready for fund4.

## Process

Steps of the process:
1. We get the Ideascale export and add columns with the Y&R cards criteria + an
"open" criteria (with a rationale). We autoflag blank assessments (`createProposerDocument.py`).
2. Proposers mark the criteria column to flag an assessment (and add a rationale for the "open" criteria).
Proposers can only flag assessments for their own proposal.
3. We make an external file with a profanity/similarity check to be used by vCA as reference (to be verified if it is possible with a great number of assessments, the similarity part is really heavy)
4. Based on the result of 1), and after the proposer have flagged the assessments, we generate a "master" file for their own review (`createVCAMaster.py`)
5. Each vCA make their own copy of the "master" file, validating a flagged review or flagging a review not previously flagged.
In this file, for each criteria there will be 2 columns (one for the proposer and one for the vCA. In this way it doesn't default to the proposer opinion. The vCA column is what will be taken in account for step 5))
Blank reviews will be also filtered out to reduce the file size.
5. We merge each single vCA file to calculate consensus based on the r/y cards document and to make an overview for the single CA (to be verified if it is possible through a script).
The count for blank assessments for each CA will be based from the file in step 1)
The script is responsible to do the math to apply the Y&R cards criteria to each assessment and to make an overview for each CA.
Final file structure:
- Assessments overview
- CA overview
- ... one sheet for each vCA review.

## Requirements Installation

```
pip3 install -r requirements.txt
```

## Usage

Create a `Service Account` for Google APIs.

1. Enable API Access for a Project if you haven’t done it yet.
2. Search for `Sheet` and `Drive` APIs and activate them.
3. Go to “APIs & Services > Credentials” and choose “Create credentials > Service account key”.
4. Fill out the form
5. Click “Create” and “Done”.
6. Press “Manage service accounts” above Service Accounts.
7. Press on `⋮` near recenlty created service account and select “Create key”.
8. Select JSON key type and press “Create”.
9. Copy the downloaded json to `gsheet-accounts/service_account.json`

Copy `options.json.template` to `options.json` and define your options.
The first time you will not have all the documents id requested.
Every script execution will create the next document that you have to add to the
options file.

1. You only have the `originalExportFromIdeascale` document.
2. Run `createProposerDocument.py` -> it will create the document where
proposers can flag assessments.
3. Insert the id of the document just created in `proposersFile`.
4. This is the proposers round. They can flag assessments in the document.
5. When they have completed the process you can run `createVCAMaster.py` that
will create the master file for VCAs.
6. Each vCA makes their own copy of the master file and assigns red/yellow cards.
7. Insert the VCA Master file id in `VCAMasterFile` options and every VCA copy
of the file in `VCAsFiles` (as a list).
8. Run `createVCAAggregate.py` to generate the final document.

### Other tools
`serviceAccountUtils.py` is used to delete every document in the service account.
Use carefully.
