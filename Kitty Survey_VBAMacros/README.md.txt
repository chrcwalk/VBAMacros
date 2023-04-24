This project was specifically created to learn and understand more about the functionality of basic and advanced macros

It has always been a goal of mine to understand these better but have not been able to do so until tackling a project such as this

Intro:
	The macros within this workbook are stored directly to the workbook and all have been manipulated using the VBA editor
	There are 4 main macros of concern - UniversalHighlight, HighlightLikelyBuyersOnlyFinalVersion, HighlightUnlikelyBuyers, and HighlightSwingBuyers
		HighlightLikelyBuyersOnlyFinalVersion - Formats all records where the answer to Q3 was "yes" and there answer to Q5 was either "Somewhat likely" or "Very likely"
		HighlightUnlikelyBuyers - Formats all records where the answer to Q3 is "No"
		HighlightSwingBuyers - Formats all records where the answer to Q3 is "Yes" but the answer to Q5 was "Unlikely"
		UniversalHighlight - This is a nested macro of all the above macro (i.e. I just combined all of the macros to one button)
	Any other macros were made for practice to get myself to the point to where I finalized the project, and all macros regardless of practice or finalized, should have explanatory notes in the code

Use Case:
	In my line of work, I've witnessed a lot of manual reports that require some amount of formatting/color coding based upon quantitative/qualitative metrics which led me to
	discover this methodology of utilizing macros through the VBA editor to push myself forward professionally and help colleagues reduce the necessary time on menial and mundane tasks

Functionality:
	HighlightLikelyBuyersOnlyFinalVersion - Will apply bolding, underlining, italicizing, and a green fill to all buyers that meet the above stated parameters
	HighlightUnlikelyBuyers - Will apply a strikethrough, bold lettering, and a red fill to all buyers that meet the above stated parameters
	HighlightSwingBuyers - Will apply bold lettering, a font size increase, and a yellow fill to all buyers that meet the above stated parameters
	UniversalHighlight - Will apply all of the above formatting rules, but with one click

Description:
	Each button was created with the utilization of loops
	There is also an option that allows for an inputbox

Important Lessons Learned:
	• VBA is a powerful internal coding tool embedded within any MS Office application and allows for far greater manipulation and automation of application functionality and internal data
	• VBA is case AND line sensitive - it matters what case the letters are and what lines each command is on (it would be nice to make that finalized "If" line of code much more readable by applying it to multiple lines, but this 			will break the code)
	• Macros and VBA editing become far easier if you break it down into individual steps and adjust as you go - Technically-minded problem solving!
	• Do While/Do Until loops are excel wizard magic sorcery tools that allow you to perform manual tasks automatically; you can embed if statements within these loops to apply conditional parameters to the loop
	• It is extremely important to understand which style of macro you are working with in terms of relative vs. non-relative macros - you can combine the two, but they must be created as separate entities and brought together later 		to keep the proper references for each respective line of code
		○ To explain a different manner: Some macros are set to "relative references" (similar to excel functions with the $ sign, which forces non-relativity) upon creation, and some are not - you ultimately choose as you create 			a new macro. If you need a combination of some non-relative references for an action, and then relative references for others, create two separate macros to perform both desired actions and then join them together 				after they have been independently created
				§ I'm hoping I'm correct about that - still learning!
You can provide user interactivity with UserForms and other options provided from the "Insert" drop-down menu