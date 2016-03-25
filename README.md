# Rubric Grader Task Pane Add-in Sample for OneNote Online (Preview)

_Applies to: OneNote Online_

This sample shows you how to use the OneNote JavaScript API (in Preview) in a OneNote task pane add-in.

The sample uses the OneNote JavaScript API to get page content, add an outline to the page, and open a different page. The add-in helps teachers to grade writing assignments based on a grading rubric.

![Rubric Grader task pane add-in in OneNote Online](readme-images/rubric-grader.png) 

## Prerequisites
- A test notebook from the OneNote team. See ///link to topic/// for more information about developing during this initial preview period.

- Microsoft Office Project Generator (Yo Office) and it's prerequisites (npm). This makes it easy to host a gulp static server. Installation instructions are in [this article](https://dev.office.com/blogs/creating-office-add-ins-with-any-editor-introducing-yo-office). 

   >You'll need Yo Office to follow the instructions in this article, but you can download the project and host your website in other environments.

## Yo Office

### Start the gulp server and create the add-in project
1. Set up the self-hosted HTTPS static web server.

1. Create a local folder named *OneNote-Addin*.

1. Create your add-in project files. For this walkthrough:
   - Choose a Task pane add-in. 
   - Choose any of the Office products.
   - Choose **HTML, CSS & JavaScript**.

1. Open the manifest-onenote-addin.xml file. Change the <SourceLocation> element of the manifest file from  **.../home/home.html* to **.../home/grader.html**.

1. After the project has been created, delete or rename the **app** folder in the **OneNote-Addin** folder. We'll be using the **app** folder from the GitHub sample.

### Configure the add-in
1. Clone or download this repo. 


### Test and run the add-in 

Trust localhost

## Learn more

