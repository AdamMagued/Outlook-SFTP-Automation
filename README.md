# Outlook-SFTP-Automation

A Java application that integrates with Microsoft Outlook to automatically process specific incoming emails and forward attachments to a secure server.

## Overview
Outlook-SFTP-Automation is a tool designed for enterprise environments to automate secure file forwarding. It monitors a specific Outlook inbox, extracts connection details (path, username, password) from the email body, downloads attachments, and uploads them to a remote SFTP server.

## Features
* **Real-time Monitoring**: Polls the Outlook inbox every 10 seconds.
* **Smart Filtering**: Filters emails by specific subjects (e.g., `SFTP-DEFAULT_HOST`).
* **Automated Parsing**: Extracts SFTP credentials and paths directly from the email body.
* **Secure Transfer**: Automatically uploads attachments to the specified remote SFTP server.

## How to Run

### Prerequisites
* Java Development Kit (JDK)
* Microsoft Outlook (configured on the host machine)

### Configuration
1.  Open the source code.
2.  Locate the main configuration variables.
3.  Replace the `String DEFAULT_HOST` and `int DEFAULT_PORT` with your specific **Linux Server IP** and **Port Number**.

### Execution
1.  Open the project in your IDE (e.g., Eclipse, IntelliJ).
2.  Locate the `class1.launch` file.
3.  Run the launch file to start the automation service.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
