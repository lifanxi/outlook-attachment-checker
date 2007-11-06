
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Attachment checking tool for Outlook
 
        November 6, 2007

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Have you ever met a problem that you send out an email saying that you attached something in the mail but actually you forgot to attach the file? This script will help you to solve this problem.

============
Installation
============
1. Open the "code.txt" file. Copy (Ctrl+C or Edit/Copy) all the text.
2. Open Outlook, go to "Tools->Macro->Visual Basice Editor".
3. On the left, you see a little tree diagram. Click the Project1/Microsoft Outlook Object/ThisOutlookSession one, then paste (CTRL+V or Edit/Paste) the code into the blank window that appears.
4. Save the code and Exit Visual Basic Editor.
5. Go to the Outlook menu "Tools->Macro->Security", change the security level from "High" to "Medium".
6. Exit Outlook and restart it. When the "Security Warning" dialog box comes out, click "Enable Macros".

=====
Usage
=====

Every time you send a email with the keyword "attach*" (case insensitive) or "¸½¼þ" (This is Chinese for the word "attachment") in the message body or subject without attaching an attachment, a message box will pop up to warn you about this and let you choose whether you really want to send the mail.

=========
Copyright
=========

Attachement checking tool for Outlook is free software: you can redistribute it and/or modify it under the terms of the GNU Lesser General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser General Public License for more details.

You should have received a copy of the GNU Lesser General Public License along with Foobar.  If not, see <http://www.gnu.org/licenses/>.


