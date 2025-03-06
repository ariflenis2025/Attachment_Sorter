import React, { useEffect, useState } from "react";
import {
  Box,
  Typography,
  Button,
  List,
  ListItem,
  ListItemIcon,
  ListItemText,
  CircularProgress,
  Snackbar,
  Alert,
  ListItemButton,
  Checkbox,
  Divider,
  Drawer,
  TextField,
  IconButton,
} from "@mui/material";
import { FaFile, FaFileImage, FaFilePdf, FaFileWord, FaFileExcel, FaFilePowerpoint, FaDatabase, FaFileAlt, FaFileArchive, FaFileAudio, FaFileCode, FaFileCsv, FaFileVideo } from "react-icons/fa";
import { RiMailSendLine, RiSettings5Line } from "react-icons/ri";
import LoaderApp from "../../Loader/Loader";

interface AppProps {
  selectdItemFromAdrees: any;
}

interface Attachment {
  id: string;
  name: string;
  contentType: string;
  contentBytes?: string;
}

const SendMailTab: React.FC<AppProps> = ({ selectdItemFromAdrees }) => {
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [selectedAttachments, setSelectedAttachments] = useState<string[]>([]);
  const [selectedEmails, setSelectedEmails] = useState<string[]>([]);
  const [isSending, setIsSending] = useState(false);
  const [loading, setLoading] = useState(false);
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [toast, setToast] = useState({
    open: false,
    severity: "info",
    message: "",
  });
  const [emailSettings, setEmailSettings] = useState({
    email1: "it@gmail.com",
    email2: "shahzad890.it@outlook.com",
    email3: "bushraiqbal.it01@gmail.com",
  });
  const [editedEmails, setEditedEmails] = useState(emailSettings);

  useEffect(() => {
    const fetchAttachments = async () => {
      setLoading(true);
      const emailAttachments = Office.context.mailbox.item.attachments || [];
      const preparedAttachments = await prepareAttachments(emailAttachments);
      setAttachments(preparedAttachments);
      setLoading(false);
    };
    const savedEmailSettings = localStorage.getItem('emailSettings');
    if (savedEmailSettings) {
      setEmailSettings(JSON.parse(savedEmailSettings));
      setEditedEmails(JSON.parse(savedEmailSettings));
    }
    fetchAttachments();
    console.log('done');
    
  }, [selectdItemFromAdrees]);

  useEffect(() => {
    // Retrieve email settings from local storage when the component mounts
    const savedEmailSettings = localStorage.getItem('emailSettings');
    if (savedEmailSettings) {
      setEmailSettings(JSON.parse(savedEmailSettings));
      setEditedEmails(JSON.parse(savedEmailSettings));
    }
  }, []);

  const handleSendEmail = () => {
    const Token = localStorage.getItem('Token');
    selectedEmails.forEach((email) => {
      sendEmailWithAttachments(email, Token, ((data, error) => {
        if (data) {
          setSelectedAttachments([]);
          setToast({
            open: true,
            severity: "success",
            message: `Email sent successfully to ${email}!`,
          });
        }
        if (error) {
          if (error.code === 'InvalidAuthenticationToken') {
            setToast({
              open: true,
              severity: "error",
              message: "Token Expired Please Login!",
            });
            LoginAgain(email);
          } else {
            setToast({
              open: true,
              severity: "error",
              message: error.message,
            });
          }
        }
      }));
    });
  };

  const LoginAgain=(email:any)=>{

    Office.onReady(() => {
  
      const dialogOptions: Office.DialogOptions = {
          height: 40,
          width: 35,
          displayInIframe: false
      }
      const redirect_uri_For_Local="https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=e5a4342f-c8a5-4185-948d-2e3d485b4822&response_type=token&redirect_uri=https://localhost:3000/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment"
      const redirect_uri_For_LIve="https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=aaa00dc4-7743-467e-8868-596ffff59e05&response_type=token&redirect_uri=https://shahzadumar-w.github.io/OutlookAddin_AttachmentSorter/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment";
  

      Office.context.ui.displayDialogAsync(
        redirect_uri_For_Local,dialogOptions,
          (asyncResult) => {
  
              if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                  console.log("Opening Dialogue Box FAILED Async" + asyncResult.error.message);
                  return;
              }
  
              const myDialog: Office.Dialog = asyncResult.value;
  
              myDialog.addEventHandler(Office.EventType.DialogMessageReceived,
                  (args: Office.DialogParentMessageReceivedEventArgs) => {
                      const token = args.message
  
                      if (typeof (token) == 'string' && token.length > 50) {
                          localStorage.setItem("Token", token);
                          sendEmailWithAttachments(email, token, ((data, error) => {
                            if (data) {
                              setToast({
                                open: true,
                                severity: "success",
                                message: "Email sent successfully!",
                              });
                              setSelectedAttachments([]);
                            }
                            if (error && error.status === 401) {
                              setToast({
                                open: true,
                                severity: "success",
                                message: "Token is Expire Please Login.!",
                              });
                              LoginAgain(email);
                            }
                          }));
                          localStorage.setItem('Token', token);
                      }
                      else {
  
                          //You Might had passed Error parameters. Debug
                          localStorage.removeItem("Token");
                          console.error("Got Empty/Incorrect response value @ Dialog window ");
                          //Debugger here (Check your value & Format. Is length fine?)
  
                      }
  
                      myDialog.close() //Closes it even incase wrong value
  
                      let existingNotifcaiton:any = Office.context.mailbox.item.notificationMessages.getAllAsync(() => { })
                      //Remove if same tag exist; no spam
  
                      Office.context.mailbox.item.notificationMessages.addAsync('Token Notifcation Success ' + existingNotifcaiton.value.length,
                          {
                              type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                              message: 'Updated Authentication',
                              icon: "icon_01",
                              persistent: false
  
                          },
  
                          () => { })
  
                  });
  
              myDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg:any) => {
  
                  console.error("Dialog Event" + arg.error); //Another Handler if dialog crashses during session time (ie window suddenly closed, access is invalid & etc
              });
  
          }
      );
  
  
  })
  }

  const prepareAttachments = async (attachments: Office.AttachmentDetails[]) => {
    return Promise.all(
      attachments.map(async (attachment) => {
        try {
          const contentBytes = await fetchAttachmentContent(attachment.id);
          return {
            id: attachment.id,
            name: attachment.name,
            contentType: attachment.contentType,
            contentBytes,
          };
        } catch (error) {
          console.error("Error preparing attachment:", attachment.name, error);
          return { id: attachment.id, name: attachment.name, contentType: attachment.contentType };
        }
      })
    );
  };

  const fetchAttachmentContent = async (attachmentId: string): Promise<string> => {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value.content);
        } else {
          console.error("Failed to fetch attachment content:", result.error.message);
          reject(result.error.message);
        }
      });
    });
  };

  const toggleAttachmentSelection = (id: string) => {
    setSelectedAttachments((prev) =>
      prev.includes(id) ? prev.filter((attachmentId) => attachmentId !== id) : [...prev, id]
    );
  };

  const toggleEmailSelection = (email: string) => {
    setSelectedEmails((prev) =>
      prev.includes(email) ? prev.filter((e) => e !== email) : [...prev, email]
    );
  };

//   const sendEmailWithAttachments = async (email: string, Token: string, callback: (data: any, error: any) => void) => {
//     try {
//       setIsSending(true);
//       const selectedAttachmentData = attachments
//         .filter((attachment) => selectedAttachments.includes(attachment.id))
//         .map((attachment) => ({
//           "@odata.type": "#microsoft.graph.fileAttachment",
//           name: attachment.name,
//           contentBytes: attachment.contentBytes,
//         }));

//       const emailData = {
//         message: {
//           subject: 'Headers me proper values set karein',
//           body: {
//             contentType: "Text",
//             // content: selectedEmails.toString(),
//             content:`SPF, DKIM, aur DMARC settings check karein.
// 2️⃣ Headers me proper values set karein.
// 3️⃣ Email content ko optimize karein (avoid spammy words).
// 4️⃣ Microsoft Defender ke spam filter rules check karein.
// 5️⃣ Bheji gai email ko manually "Not Junk" mark karein, taake AI model seekh sake.'`
//           },
//           toRecipients: [
//             {
//               emailAddress: {
//                 address: email,
//               },
//             },
//           ],
//           attachments: selectedAttachmentData,
//         },
//         saveToSentItems: "true",
//       };

//       const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
//         method: "POST",
//         headers: {
//           "Content-Type": "application/json",
//           Authorization: `Bearer ${Token}`,
//         },
//         body: JSON.stringify(emailData),
//       });

//       if (response.ok) {
//         setToast({
//           open: true,
//           severity: "success",
//           message: "Email sent successfully!",
//         });
//         callback('Email sent successfully', null);
//       } else {
//         const errorData = await response.json();
//         setToast({
//           open: true,
//           severity: "error",
//           message: errorData.error.message || "Failed to send email.",
//         });
//         callback(null, errorData.error);
//       }
//     } catch (error) {
//       setToast({
//         open: true,
//         severity: "error",
//         message: `Error: ${error.message}`,
//       });
//       callback(null, error);
//     } finally {
//       setIsSending(false);
//     }
//   };


const sendEmailWithAttachments = async (
  email: string, 
  Token: string, 
  callback: (data: any, error: any) => void
) => {
  try {
    setIsSending(true);

    // Ensure the email has a subject and meaningful body
    const emailData = {
      message: {
        subject: "Receipt", // 🔹 Must have a subject
        body: {
          contentType: "Text",
          content:  selectedEmails.toString()// 🔹 Must have a body
        },
        toRecipients: [{ emailAddress: { address: email } }],
        attachments: attachments
          .filter(att => selectedAttachments.includes(att.id))
          .map(att => ({
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: att.name,
            contentBytes: att.contentBytes
          }))
      },
      saveToSentItems: "true"
    };

    const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${Token}`,
      },
      body: JSON.stringify(emailData),
    });

    if (response.ok) {
      setToast({ open: true, severity: "success", message: "Email sent successfully!" });
      callback("Email sent successfully", null);
    } else {
      const errorData = await response.json();
      setToast({ open: true, severity: "error", message: errorData.error.message || "Failed to send email." });
      callback(null, errorData.error);
    }
  } catch (error) {
    setToast({ open: true, severity: "error", message: `Error: ${error.message}` });
    callback(null, error);
  } finally {
    setIsSending(false);
  }
};

  const handleDrawerSave = () => {
    setEmailSettings(editedEmails);
    localStorage.setItem('emailSettings', JSON.stringify(editedEmails)); // Save to local storage
    setDrawerOpen(false);
  };

  const handleCancel = () => {
    setDrawerOpen(false);
  };

  const getIcon = (type: string) => {
    switch (type) {
      case "image/png":
      case "image/jpeg":
      case "image/jpg":
      case "image/gif":
      case "image/webp":
      case "image/svg+xml":
      case "image/bmp":
        return <FaFileImage color="blue" fontSize={'xx-large'} />;
      case "application/pdf":
        return <FaFilePdf color="red" fontSize={'xx-large'} />;
      case "application/msword":
      case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        return <FaFileWord color="blue" fontSize={'xx-large'} />;
      case "application/vnd.ms-excel":
      case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
      case "text/csv":
        return <FaFileExcel color="green" fontSize={'xx-large'} />;
      case "text/csv":
        return <FaFileCsv color="teal" fontSize={'xx-large'} />;
      case "application/vnd.ms-powerpoint":
      case "application/vnd.openxmlformats-officedocument.presentationml.presentation":
        return <FaFilePowerpoint color="orange" fontSize={'xx-large'} />;
      case "application/zip":
      case "application/x-7z-compressed":
      case "application/x-rar-compressed":
      case "application/x-tar":
      case "application/gzip":
        return <FaFileArchive color="brown" fontSize={'xx-large'} />;
      case "text/plain":
        return <FaFileAlt color="gray" fontSize={'xx-large'} />;
      case "audio/mpeg":
      case "audio/wav":
      case "audio/ogg":
      case "audio/aac":
      case "audio/x-midi":
        return <FaFileAudio color="purple" fontSize={'xx-large'} />;
      case "video/mp4":
      case "video/x-msvideo":
      case "video/mpeg":
      case "video/ogg":
      case "video/webm":
      case "video/x-matroska":
        return <FaFileVideo color="red" fontSize={'xx-large'} />;
      case "text/html":
      case "application/javascript":
      case "application/json":
      case "text/css":
      case "application/xml":
        return <FaFileCode color="green" fontSize={'xx-large'} />;
      case "application/vnd.ms-access":
      case "application/x-sqlite3":
      case "application/x-msdownload":
        return <FaDatabase color="brown" fontSize={'xx-large'} />;
      default:
        return <FaFile color="black" fontSize={'xx-large'} />;
    }
  };
  return (
    <Box sx={{ maxWidth: 600, margin: "auto", backgroundColor: "#fff" }}>
      {loading && <LoaderApp />}
      {isSending && <LoaderApp />}

      <Typography variant="h5" sx={{ marginBottom: 2, textAlign: "center", fontWeight: "bold", display: 'flex', alignItems: 'center', justifyContent: 'space-between', fontSize: '20px' }}>
        <div>
          <RiMailSendLine style={{ marginRight: 8, verticalAlign: "middle" }} />
          Send Email
        </div>
        <IconButton
          sx={{ float: "right", marginBottom: '5px' }}
          color="primary"
          onClick={() => setDrawerOpen(true)}
        >
          <RiSettings5Line />
        </IconButton>
      </Typography>

      <Box sx={{ maxHeight: 300, overflow: "auto" }}>
        <List>
          {attachments.length === 0 && !loading && (
            <Box textAlign="center" p={2}>
              No Attachment Found with the selected mail
            </Box>
          )}
          {attachments.map((attachment) => (
            <React.Fragment key={attachment.id}>
              <ListItem disablePadding>
                <Checkbox
                  checked={selectedAttachments.includes(attachment.id)}
                  onChange={() => toggleAttachmentSelection(attachment.id)}
                />
                <ListItemButton>
                  <ListItemIcon>{getIcon(attachment.contentType)}</ListItemIcon>
                  <ListItemText sx={{ fontSize: '14px' }} primary={attachment.name} />
                </ListItemButton>
              </ListItem>
              <Divider />
            </React.Fragment>
          ))}
        </List>
      </Box>

      <Typography variant="h6" sx={{ marginTop: 2, fontSize: "15px" }}>
        Select Email Addresses
      </Typography>
      <List>
        {Object.entries(emailSettings).map(([key, email]) => (
          <ListItem key={key} disablePadding>
            <Checkbox
              checked={selectedEmails.includes(email)}
              onChange={() => toggleEmailSelection(email)}
            />
            <ListItemText primary={email} />
          </ListItem>
        ))}
      </List>

      <Button
        variant="contained"
        color="primary"
        fullWidth
        sx={{
          "&.Mui-disabled": {
            backgroundColor: "#93C0FE",
            color: "white"
          },
          marginTop: 2,
        }}
        disabled={isSending || !selectedAttachments.length || !selectedEmails.length}
        onClick={handleSendEmail}
        startIcon={<RiMailSendLine />}
      >
        Send to Selected Emails
      </Button>

      <Snackbar
        open={toast.open}
        autoHideDuration={50000}
        onClose={() => setToast({ ...toast, open: false })}
        sx={{
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center'
        }}
      >
        <Alert severity={toast.severity as any} onClose={() => setToast({ ...toast, open: false })}>
          {toast.message}
        </Alert>
      </Snackbar>

      <Drawer anchor="right" open={drawerOpen} onClose={() => setDrawerOpen(false)}>
        <Box sx={{ padding: 2 }}>
          <Typography variant="h6" sx={{ marginBottom: 2 }}>
            Edit Email Addresses
          </Typography>
          {Object.entries(editedEmails).map(([key, email]) => (
            <TextField
              key={key}
              fullWidth
              label={`Email ${key.split("email")[1]}`}
              value={email}
              onChange={(e) => setEditedEmails((prev) => ({ ...prev, [key]: e.target.value }))}
              sx={{ marginBottom: 2 }}
            />
          ))}
          <div style={{ width: '100%', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <Button variant="outlined" sx={{ width: '49%' }} onClick={handleCancel}>
              Cancel
            </Button>
            <Button variant="contained" color='primary' sx={{ width: '49%' }} onClick={handleDrawerSave}>
              Save
            </Button>
          </div>
        </Box>
      </Drawer>
    </Box>
  );
};

export default SendMailTab;