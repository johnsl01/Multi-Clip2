# Multi-Clip2

Simple Multiple Clipboard Utility

Implements a basic text and image 5 pane clipboard.  

Doesn't support complex clipboard contents - so using it for someting cut from Excel or Word (for examples) will loose all formatting but will keep the text content (like pasting into notepad).

Each of the 5 clipboards can hold both an image and text data at the same time - and can be switched between image and text mode - any actions on the image will not affect the text - any any action on the text will not affect the image.

Each tiny pane supports text content editing - but their primary purpose is to show the content of the text in the clipboard.

In image mode the pane shows a resized image of the content - the actual content is not resized - so again it is just to remind the user what is in each clipboard.

A useful enhancement would be to save and restore sets of content - would be handy to preload with up to 5 commonly used strings for repetitave actions, for example.

And a further more complex enhancement would be to store different clipboard content types so Word, Excel or other complex content can be pasted in without loosing the formatting and other non-text content.
