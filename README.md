# Submit an Microsoft Office document example

Future version of officegen will support also importing existing Word (docx), PowerPoint (pptx) and Excell (xlsx) files. In addition to that it's much easier to demonstrate your requested feature using an example file.
To support this feature I'm asking the community to submit example files as a PR. To do that you MUST follow the following instructions:

- [ ] Your file must demonstrate a feature or a tool/version that created this file (for example: Office 2019 for Mac).
- [ ] Make sure that your file is not too big. Few MB max.
- [ ] No sensitive, illegal, secret, copyrighted stuff!
- [ ] Please check if someone already posted similar file in the files-for-parsing branch.
- [ ] Clone the files-for-parsing branch:

```
git clone --single-branch --branch files-for-parsing https://github.com/Ziv-Barber/officegen.git
```

- [ ] Create a directory to hold your files:
	- project-root-directory/pptx_files/your-github-user = for Powerpoint files.
	- project-root-directory/docx_files/your-github-user = for Word files.
	- project-root-directory/xlsx_files/your-github-user = for Excel files.
- [ ] For each file you need to create 2 files:
	- [ ] The example file itself. Name it related to the feature that you requesting.
	- [ ] Create .md file with the SAME base name as the example file. Use the template:
		- project-root-directory/_templates/pptx_files/x_files/example.md = for Powerpoint files.
		- project-root-directory/_templates/docx_files/x_files/example.md = for Word files.
		- project-root-directory/_templates/xlsx_files/x_files/example.md = for Excel files.
- [ ] Don't change anything else outside your github user directory under either pptx_files, docx_files or xlsx_files!
- [ ] Click [here](https://github.com/Ziv-Barber/officegen/compare/files-for-parsing...your-branch?expand=1&template=PR_SUBMIT_FILE.md) to create a PR.

Thanks!
