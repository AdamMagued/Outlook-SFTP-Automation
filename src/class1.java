import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.jcraft.jsch.*;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

public class class1 {

    // Local folder to save attachments
    private static final String DOWNLOAD_DIR = "D:\\Attachments\\";

    // SFTP default values (host and port)
    private static final String DEFAULT_HOST = ""; // ENTER THE LINUX SERVER IP HERE
    private static final int DEFAULT_PORT = ; // ENTER PORT HERE

    public static void main(String[] args) {
        ensureDirectory(DOWNLOAD_DIR);

        while (true) {
            try {
                processInbox();
                System.out.println("Checked inbox at " + new java.util.Date());
                Thread.sleep(10000); // wait 10 sec
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    private static void processInbox() {
        try {
            ActiveXComponent outlook = new ActiveXComponent("Outlook.Application");
            Dispatch mapiNamespace = outlook.getProperty("Session").toDispatch();
            Dispatch inbox = Dispatch.call(mapiNamespace, "GetDefaultFolder", new Variant(6)).toDispatch();
            Dispatch items = Dispatch.get(inbox, "Items").toDispatch();

            int count = Dispatch.get(items, "Count").getInt();

            for (int i = 1; i <= count; i++) {
                Dispatch mailItem = Dispatch.call(items, "Item", i).toDispatch();
                String subject = Dispatch.get(mailItem, "Subject").toString();
                Variant unreadVariant = Dispatch.get(mailItem, "UnRead");
                boolean isUnread = unreadVariant.getBoolean();

                if (isUnread && subject.contains("SFTP-10.99.161.41")) {
                    System.out.println("Found matching email: " + subject);

                    // Get plain-text body only
                    String body = strOrNull(Dispatch.get(mailItem, "Body"));

                    // Debug dump
                    System.out.println("----- RAW BODY START -----");
                    if (body == null) {
                        System.out.println("(null)");
                    } else {
                        System.out.println(body.replace("\r", "\\r").replace("\n", "\\n\n"));
                    }
                    System.out.println("----- RAW BODY END -----");

                    Creds creds = parseCredentials(body);
                    if (creds == null) {
                        System.out.println("Could not parse body, skipping this email.");
                        continue;
                    }

                    // Get attachments
                    Dispatch attachments = Dispatch.get(mailItem, "Attachments").toDispatch();
                    int attCount = Dispatch.get(attachments, "Count").getInt();

                    for (int j = 1; j <= attCount; j++) {
                        Dispatch attachment = Dispatch.call(attachments, "Item", j).toDispatch();
                        String fileName = Dispatch.get(attachment, "FileName").toString();
                        String localPath = DOWNLOAD_DIR + fileName;

                        // Save attachment locally
                        Dispatch.call(attachment, "SaveAsFile", localPath);
                        System.out.println("Saved: " + localPath);

                        // Upload via SFTP using parsed creds
                        uploadFile(localPath, creds);
                    }

                    // Mark email as read
                    Dispatch.put(mailItem, "UnRead", false);
                    System.out.println("Done processing email: " + subject);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static Creds parseCredentials(String body) {
        if (body == null) return null;

        String[] rawLines = body.split("\\r?\\n");
        List<String> filtered = new ArrayList<>();
        for (String l : rawLines) {
            if (l != null && !l.trim().isEmpty()) {
                filtered.add(l.trim());
            }
        }

        if (filtered.size() < 3) {
            return null;
        }

        String remotePath = filtered.get(0);
        String user = filtered.get(1);
        String pass = filtered.get(2);

        System.out.println("Parsed from body:");
        System.out.println("Path: " + remotePath);
        System.out.println("User: " + user);
        System.out.println("Pass: " + pass);

        return new Creds(remotePath, user, pass);
    }

    private static void uploadFile(String localFilePath, Creds creds) {
        Session session = null;
        Channel channel = null;

        try {
            JSch jsch = new JSch();
            session = jsch.getSession(creds.user, DEFAULT_HOST, DEFAULT_PORT);
            session.setPassword(creds.pass);

            Properties config = new Properties();
            config.put("StrictHostKeyChecking", "no");
            session.setConfig(config);

            session.connect();
            channel = session.openChannel("sftp");
            channel.connect();

            ChannelSftp sftp = (ChannelSftp) channel;
            sftp.put(localFilePath, creds.remotePath);

            System.out.println("Uploaded: " + localFilePath + " -> " + creds.remotePath);

            sftp.exit();
        } catch (Exception e) {
            System.err.println("Upload failed for " + localFilePath);
            e.printStackTrace();
        } finally {
            if (channel != null && channel.isConnected()) channel.disconnect();
            if (session != null && session.isConnected()) session.disconnect();
        }
    }

    private static void ensureDirectory(String dirPath) {
        File dir = new File(dirPath);
        if (!dir.exists()) {
            if (dir.mkdirs()) {
                System.out.println("Created directory: " + dirPath);
            }
        }
    }

    private static String strOrNull(Variant v) {
        if (v == null) return null;
        String s = v.toString();
        return (s == null || s.equals("null")) ? null : s;
    }

    private static class Creds {
        String remotePath;
        String user;
        String pass;
        Creds(String r, String u, String p) {
            this.remotePath = r;
            this.user = u;
            this.pass = p;
        }
    }
}



