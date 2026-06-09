# AppScripts Code Blocks

1. Initial Setup
Enable APIs: In your Apps Script project, add both the Drive API and the Google Docs API.
<img width="748" height="1013" alt="image" src="https://github.com/user-attachments/assets/481472be-92ab-4204-bfbe-74278bd16b7e" />

Add the Script: Open your existing script file, delete any default template code, and paste the contents of code_blocks_code.gs. Save the file.

Create the Sidebar: Create a new HTML file, name it sidebar, and paste the contents of code_blocks_Sidebar.html.
<img width="330" height="159" alt="image" src="https://github.com/user-attachments/assets/e320f9d9-91a9-4273-bcd2-5960d03495b0" />

Save & Refresh: Save your project, return to your Google Doc, and refresh the page.

2. How to Use the Formatter
Click Code Formatter in the Google Docs menu.
<img width="651" height="51" alt="image" src="https://github.com/user-attachments/assets/3d4f0ef1-49d2-48ef-9d37-bcd7c2fa4020" />

Grant the necessary permissions when prompted.

Highlight the text you want to format.

Select your programming language in the sidebar.

Review the preview, then click Insert Code Block into Doc.
<img width="324" height="418" alt="image" src="https://github.com/user-attachments/assets/c951fe5f-cc4a-48ab-8c1b-ef01cc74ed18" />

⚠️ Known Issue: Re-formatting an already formatted block may duplicate the line numbers. To fix this, either manually delete the extra numbers, undo the change using Cmd/Ctrl + Z, or paste the original unformatted code and try again.
