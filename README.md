# Outlook-SFTP-Automation
Replace String DEFAULT_HOST and  int DEFAULT_PORT with your LINUX SERVER IP and Port number
To run the code start class1.launch file

## Overview
Outlook-SFTP-Automation is a Java application that integrates with Microsoft Outlook to automatically process specific incoming emails.  
It downloads attachments, extracts connection details (path, username, password) from the email body, and uploads the attachments to a remote SFTP server.  

This tool is useful for automating secure file forwarding from Outlook to SFTP in enterprise environments.
- Monitors Outlook inbox in real-time (polling every 10 seconds).  
- Filters emails by subject (e.g. `SFTP-DEFAULT_HOST`).  
- Parses SFTP connection details from the **email body**:  
