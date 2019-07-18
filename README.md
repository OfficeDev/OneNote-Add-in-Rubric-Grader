---
topic: sample
products:
- office-onenote
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 3/24/2016 8:29:25 AM
---
# Rubric Grader Task Pane Add-in for OneNote Online

_Applies to: OneNote Online_

The Rubric Grader sample shows you how to use the OneNote JavaScript API in a OneNote task pane add-in. The add-in gets page content, adds an outline to the page, and opens a different page.

The add-in helps teachers grade writing assignments based on a grading rubric.

![Rubric Grader task pane add-in in OneNote Online](readme-images/rubric-grader.png) 

## Prerequisites
- Yeoman Office generator. To install the generator and its prerequisites, follow these [installation instructions](https://dev.office.com/docs/add-ins/get-started/create-an-office-add-in-using-any-editor). After you install the Yeoman Office generator by using the npm command, return to this readme. 

   The Yeoman Office generator makes it easy to create add-in projects when you don't have Visual Studio or you want to use technologies other than plain HTML, CSS, and JavaScript. It also provides quick access to a local Gulp web server for testing. 

   >You can optionally [use Visual Studio](https://dev.office.com/docs/add-ins/get-started/create-and-debug-office-add-ins-in-visual-studio) to create a task pane add-in project, and then host the sample's **app** folder content on any HTTPS website. If you use this method, skip Step 2 below, and point the **SourceLocation** in the manifest to **grader.html** on your website in Step 3.

## Step 1: Download the sample
1. Clone or download the [OneNote-Add-in-Rubric-Grader](https://github.com/OfficeDev/oneNote-Add-in-Rubric-Grader) repository. 

   The Office Add-in Generator creates a lot of supporting files for add-in projects. Most of these files aren't stored in the repository, so you'll generate a local project and then overwrite some local files with sample files. 

## Step 2: Create the add-in project and set up the test server
1. Create a local folder named *onenote add-in*.

2. Open a **cmd** prompt and navigate to the **onenote add-in** folder. Run the `yo office` command as shown below.

   ```
C:\your-local-path\onenote add-in\> yo office
   ```
   >These instructions use the Windows command prompt, but are equally applicable for other shell environments. 

3. Use the following options to create the project.

   | Option | Value |
   |:------|:------|
   | Project name | OneNote Add-in |
   | Root folder of project | (accept the default) |
   | Office project type | Task Pane Add-in |
   | Supported Office applications | (choose any--we'll add a OneNote host later) |
   | Technology to use | HTML, CSS & JavaScript |

   It takes a few minutes to create the project and add all the supporting files.

4. After the project is created, run `gulp serve-static` in the **cmd** prompt as shown below. This will start the Gulp web server.

   ```
C:\your-local-path\onenote add-in\> gulp serve-static
   ```
   The server is available when you see the `Finished 'serve-static' ...` entry in the window. Keep this window open while you're running the add-in.

5. Install the Gulp web server's self-signed certificate as a trusted certificate. You only need to do this one time on your computer for add-in projects created with the Yeoman Office generator.  

   a. In a browser, navigate to the hosted add-in page. By default, this is the same URL that's in your manifest:

   ```
https://localhost:8443/app/home/home.html
   ```

   b. Install the certificate as a trusted certificate. For more information, see [Adding Self-Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/docs/trust-self-signed-cert.md).

## Step 3: Configure the add-in project 
1. Open the **onenote add-in** folder that you created, and delete the **app** folder from the project files.

2. Copy the **app** folder from the sample files into your **onenote add-in** folder to replace the one you just deleted.

3. Open **manifest-onenote-add-in.xml** in the **onenote add-in** folder.

   a. Add the following line to the **Hosts** section. This specifies that your add-in supports the OneNote host application.

   ```
<Host Name="Notebook" />
   ```

   b. In the **DefaultSettings** section, change the **SourceLocation** element from  `home.html` to `grader.html`, as shown below.

   ```
<SourceLocation DefaultValue="https://localhost:8443/app/home/grader.html"/>
   ```

## Step 4: Run the add-in 
1. In OneNote Online, open a notebook that contains a couple of pages. Make sure at least one page has a paragraph of content.

   >If this is the first time you've opened OneNote Online, you may need to refresh the page to see your default notebook.

2. Choose **Insert > Office Add-ins**. This opens the Office Add-ins dialog. 
  - If you're logged in with your work or school account, choose the **MY ORGANIZATION** tab, and then choose  **Upload My Add-in**.
  - If you're logged in with your consumer account, choose the **MY ADD-INS** tab, and then choose  **Upload My Add-in**. (In some cases, you might need to choose **Manage My Add-ins > Upload My Add-in**.)
  

  >To enable the **Office Add-ins** button, click inside the OneNote page.

3. In the Upload Add-in dialog, browse to **manifest-onenote-add-in.xml** that you modified in your project files, and then choose **Upload**. While testing, your manifest file can be stored locally.

4. The add-in opens in an iFrame next to the OneNote page. You can:
   - Use the **Get count** button to get approximate word and sentence counts. 
   - Set scores in the scoring dropdowns, enter a comment, and then choose **Grade it** to add the grade to the page.
   - Choose **Open page** to open the page that's selected in the dropdown.

### Troubleshooting and tips 
- You can debug the add-in using your browser's developer tools. When you're using the Gulp web server and debugging in Internet Explorer or Chrome, you can save your changes locally and then just refresh the add-in's iFrame.

- When you inspect a OneNote object, the properties that are currently available for use display actual values. Properties that need to be loaded display *undefined*. Expand the `__proto__` node to see properties that are defined on the object, but are not yet loaded.

      ![Unloaded OneNote object in the debugger](readme-images/debug.png)

- You need to enable mixed content in the browser if your add-in uses any HTTP resources. Production add-ins should use only secure HTTPS resources.

## Questions and comments
We'd love to get your feedback about this sample. You can send your questions and suggestions to us in the [Issues](https://github.com/OfficeDev/oneNote-Add-in-Rubric-Grader/issues) section of this repository.

Questions about OneNote development in general should be posted to [Stack Overflow](http://stackoverflow.com/questions/tagged/onenote-api). Make sure that your questions or comments are tagged with *[onenote-api]*.

You can suggest changes for the OneNote JavaScript API on [UserVoice](https://onenote.uservoice.com/forums/245490-onenote-apis/filters/top).
  
## Additional resources

- [OneNote add-ins JavaScript programming overview](http://aka.ms/onenote-add-ins)
- [Office Add-ins platform overview](https://dev.office.com/docs/add-ins/overview/office-add-ins)

## Copyright
Copyright (c) 2016 Microsoft. All rights reserved.


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
