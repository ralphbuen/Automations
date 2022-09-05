# Data Encoding Tool
Tool made for Data Inputs in Toolbox. These tools are created using Microsoft Excel VBA.
## Tool Creation
###### This part will show you how to create a tool from scratch
1. First, you have to go to the update page. To do that, you have to click on:
  - Management
    - Maintenance
      - Product Hub
        - Updates
          - Field Updates

2. After navigating to the page, select the lender that you want to make an automation tool for.

3. After selecting a lender, decide if it will be for Interest Rate - Initial Rate or Interest Rate - 1st Revert Rate. Then click search

4. When the page has fully loaded, press Ctrl + S on the keyboard and save the webpage anywhere you see fit

5. After saving the webpage, open a new Excel file and save it as an 'Excel Macro-Enabled Worksheet' (example file name would be: Westpac Automation.xlsm)

6. Open the HTML file and you should be presented with the full HTML code for the front end

https://user-images.githubusercontent.com/106564201/188057598-9411b3e1-a784-4c81-a624-19d1d7e1d4fd.mp4

7. Navigate 

to the `<td colspan="1">Interest Rate – Initial Rate</td></tr>` part of the code. There you will see `<tr id="` tags that we will be keeping. 

https://user-images.githubusercontent.com/106564201/188073599-79a00ffb-4b22-4dfe-818a-cb54029edb09.mp4

8. What you will do next is to highlight from the `<!DOCTYPE html` part of the code up until the `<td colspan="1">Interest Rate – Initial Rate</td></tr>` part of the code and then delete the highlighted codes.

https://user-images.githubusercontent.com/106564201/188073648-fc3b757d-592d-455f-b7ca-6584bf96bbfa.mp4

9. After deleting the first parts of the code, navigate to the `</tbody></table>` parts of the code and delete all of the codes until the last bit.

https://user-images.githubusercontent.com/106564201/188073676-634858d3-23f1-439d-827d-a4c0570868a2.mp4

10. After deleting the unnecessary blocks of codes, you have to clean up the code for you to extract the IDs and names properly. First, we will get all of the `<tr id="`s presented per line. Using Notepad++, press `Ctrl + A` to select all of the codes and then we will be replacing `<tr id="` with `\n<tr id="` to add a new line before every `<tr id="`. This will make sure that we will get the specific items that we need from the code. **You have to select _Regular Expression_ in the Search Mode for this to work properly**

https://user-images.githubusercontent.com/106564201/188073707-2bd081cd-3e24-49e7-a6b3-8e8780966038.mp4

11. After furnishing the `<tr id="`s, we will get rid of the `</span>` tags as well as the other codes that follows it. To do that, we will press `Ctrl + A` to select all of the codes and then we will be finding `</span>.*`. This means that it's going to find the `</span>` syntax and everything that follows `</span>`. Then we will replace it with nothing which will erase everything that follows `</span>`. Don't worry about it deleting the full block of code because it's just gonna do this line by line.

https://user-images.githubusercontent.com/106564201/188075650-470000f6-c0dd-4f32-9477-c48d910d640a.mp4

12. After removing the `</span>` tags and the following codes, what you have to do is to replace the `&lt;` with the < symbol, `&gt;` with the > symbol, and `&amp;` with the & symbol

https://user-images.githubusercontent.com/106564201/188076157-f95c8ca5-f5b3-44d0-9981-43ea62551972.mp4

13. After replacing the symbols, we will now remove the syntaxes starting from `-container"` up until `title">`. To do this, we have to type `-container.*.">` in the Find section. What this does is that it's going to look for the `-container"` as a starting position - everything in the middle - up until the `">` symbol. After setting `-container.*.">` in the Find section, we will now replace it with a `\t` for a tab. We will replace it with a tab so that when we paste it in Excel, they will be separated in two different columns.

https://user-images.githubusercontent.com/106564201/188078951-64a2f607-edc0-492b-bb1b-19c024100a62.mp4

14. Now we will be removing the `<tr id="`s. To do this, we have to type `<tr id="` in the Find section. After that, we will now replace it with nothing to delete the tag.

https://user-images.githubusercontent.com/106564201/188338144-c5a4617b-469d-4105-872b-8a400d783556.mp4

15. After modifying the file, paste it on Excel and it will look like the screenshot below

![image](https://user-images.githubusercontent.com/106564201/188343248-5e544a0e-3184-437f-bd57-a20a80c20496.png)

16. Interchange the two columns in any method that you want so it would look like this:

![image](https://user-images.githubusercontent.com/106564201/188343305-1a28d84e-5a87-4711-ac82-93f2b9689000.png)

17. Now, copy and paste the modified file in the Notepad++ once again.

