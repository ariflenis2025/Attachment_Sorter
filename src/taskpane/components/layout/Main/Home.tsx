import React, { useEffect, useState } from "react";
import {
  Box,
  Tab,
  Tabs,
  Typography,
  Button,
  TextField,
  List,
  ListItem,
  ListItemButton,
  ListItemIcon,
  ListItemText,
  Divider,
  Snackbar,
  Alert,
  CircularProgress,
} from "@mui/material";
import {
  FaFile,
  FaFileImage,
  FaFilePdf,
  FaFileWord,
  FaFileExcel,
  FaFilePowerpoint,
  FaFileAlt,
  FaFileArchive,
  FaFileAudio,
  FaFileVideo,
  FaFileCode,
  FaDatabase,
  FaFileCsv,
} from "react-icons/fa";
import UploadAttachmentsToOneDrive from "./UploadAttachmentsToOneDrive";
import LoaderApp from "../../Loader/Loader";
import SendMailTab from "./SendMail";
import { useNavigate } from 'react-router-dom';
import PDFGenrator from "./PDFGenrator";

interface AppProps{
  selectdItemFromAdrees:any
}


const AttachmentSorter: React.FC <AppProps>= ({selectdItemFromAdrees}) => {
  const [value, setValue] = useState(0);
  const [isSending, setIsSending] = useState(false);
  const [attachments, setAttachments] = useState<any[]>([]);
  const [emailBody, setEmailBody] = useState("");
  const [recipientEmail, setRecipientEmail] = useState("");
  const [ValidToken, setValidToken] = useState("");
  const [Loading, setLoading] = useState(false);
  const [SelectedemailBody, setSelectedemailBody] = useState("");

  const [toast, setToast] = useState({
    open: false,
    severity: "info", // info, success, warning, error
    message: "",
  });
  const navigate = useNavigate();

  useEffect(() => {
    const checkTokenValidity = () => {
      const token = localStorage.getItem('Token');
      if(!token){
        navigate('/')
      }
      setValidToken(token)
    };

    // Check once on load
    checkTokenValidity();

    // Check every minute
    const interval = setInterval(checkTokenValidity, 60000);

    return () => clearInterval(interval); // Cleanup on component unmount
  }, [navigate]);

  useEffect(() => {
    setLoading(true)
    getSelectedEmailBody()
    const attachments = Office.context.mailbox.item.attachments || [];
    setAttachments(attachments);
    setLoading(false)
  }, []);



  useEffect(() => {
    setLoading(true)
    getSelectedEmailBody()
    const attachments = Office.context.mailbox.item.attachments || [];
    setAttachments(attachments);
    setLoading(false)
  }, [selectdItemFromAdrees]);

  // Function to determine the icon based on attachment type
 
  
  function getSelectedEmailBody() {
    // Ensure the item is a message
    const item = Office.context.mailbox.item;
  
    if (item) {
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("Email body retrieved successfully:");
          console.log(result.value); // HTML content of the email body
          setSelectedemailBody(result.value)

        } else {
          console.error("Failed to get email body:", result.error.message);
        }
      });
    } else {
      console.error("No email item is selected.");
    }
  }
 
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

  const prepareAttachments = async () => {
    const preparedAttachments = await Promise.all(
      attachments.map(async (attachment) => {
        const contentBytes = await fetchAttachmentContent(attachment.id);
        return {
          "@odata.type": "#microsoft.graph.fileAttachment",
          name: attachment.name,
          contentType: attachment.contentType,
          contentBytes,
        };
      })
    );
    return preparedAttachments;
  };

  const sendEmailWithAttachments = async (token: string) => {
    try {
      const attachmentData = await prepareAttachments();

      const emailData = {
        message: {
          subject: "Email from Attachment Sorter",
          body: {
            contentType: "Text",
            content: emailBody,
          },
          toRecipients: [
            {
              emailAddress: {
                address: recipientEmail,
              },
            },
          ],
          attachments: attachmentData,
        },
        saveToSentItems: "true",
      };

      const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${token}`,
        },
        body: JSON.stringify(emailData),
      });

      if (response.ok) {
        setIsSending(false)
    setLoading(false)

        setToast({
          open: true,
          severity: "success",
          message: "Email sent successfully!",
        });
      } else {
        const errorData = await response.json();
        console.error("Error sending email:", errorData);
    setLoading(false)

        setIsSending(false)
        setToast({
          open: true,
          severity: "error",
          message: errorData.error.message || "Unknown error occurred.",
        });
      }
    } catch (error) {
      console.error("Error:", error);
    setLoading(false)

      setToast({
        open: true,
        severity: "error",
        message: `Failed to send email: ${error}`,
      });
    }
  };

  const handleSendAttachment = async () => {
    setLoading(true)
    setIsSending(true)
    if (!recipientEmail) {
    setLoading(false)

      setIsSending(false)
      setToast({
        open: true,
        severity: "warning",
        message: "Please enter a valid recipient email address.",
      });
      return;
    }
   let storeToken=localStorage.getItem('Token')
    sendEmailWithAttachments(ValidToken);
  };

  const handleTabChange = (_event: React.SyntheticEvent, newValue: number) => {
    setValue(newValue);
  };

  const handleToastClose = (_event?: React.SyntheticEvent | Event, reason?: string) => {
    if (reason === "clickaway") return;
    setToast({ ...toast, open: false });
  };

  return (
    <Box sx={{ width: "100%" }}>
      {Loading && (<LoaderApp/>)}
      <Tabs
        TabIndicatorProps={{
          style: {
            backgroundColor: "black",
          },
        }}
        value={value}
        onChange={handleTabChange}
        aria-label="Attachment and Text Management Tabs"
        style={{ background: "#0095ff" }}
      >
        <Tab label="Send Mail" style={{ width: "33%", fontWeight: "bold", color: "white" }} />
        <Tab label="Upload" style={{ width: "33%", fontWeight: "bold", color: "white" }} />
        <Tab label="Send PDF " style={{ width: "33%", fontWeight: "bold", color: "white" }} />
        
      </Tabs>
      <Box sx={{ padding: 2 }}>
        {value === 0 && (
          <Box>
   <SendMailTab selectdItemFromAdrees={selectdItemFromAdrees} />
          </Box>
        )}

        {value === 1 && (
          <UploadAttachmentsToOneDrive  selectdItemFromAdrees={selectdItemFromAdrees}/>
        )}

        {value === 2 && (
        <PDFGenrator  selectdItemFromAdrees={selectdItemFromAdrees}/>
        )}
      </Box>
      {/* Snackbar for Toast Messages */}
      <Snackbar
        open={toast.open}
        autoHideDuration={6000}
        onClose={handleToastClose}
        anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
      >
        <Alert onClose={handleToastClose} severity={toast.severity as any} sx={{ width: "100%" }}>
          {toast.message}
        </Alert>
      </Snackbar>
    </Box>
  );
};

export default AttachmentSorter;
