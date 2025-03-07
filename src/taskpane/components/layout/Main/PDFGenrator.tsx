// import React, { useState, useEffect } from "react";
// import {
//   Box,
//   Typography,
//   TextField,
//   Button,
//   IconButton,
//   Drawer,
//   Card,
//   CardHeader,
//   CardContent,
//   Avatar,
//   CardActions,
//   Dialog,
//   DialogTitle,
//   DialogContent,
//   DialogActions,
//   Alert,
//   Snackbar,
//   Divider,
//   Select,
//   MenuItem,
//   FormControl,
//   InputLabel,
//   ListItemText,
//   Checkbox,
// } from "@mui/material";
// import { red } from "@mui/material/colors";
// import { RiMailSendLine, RiSettings5Line } from "react-icons/ri";
// import { FaArrowsLeftRight, FaRegCopy } from "react-icons/fa6";
// import pdfMake from "pdfmake/build/pdfmake";
// import pdfFonts from "pdfmake/build/vfs_fonts";
// import { OpenDialog } from "../../utils/OpenDialog";
// import axios from "axios";
// import LoaderApp from "../../Loader/Loader";

// // Ensure vfs is set properly
// if (pdfFonts && pdfFonts.pdfMake && pdfFonts.pdfMake.vfs) {
//   pdfMake.vfs = pdfFonts.pdfMake.vfs;
// } else {
//   console.log("pdfFonts.vfs is not available.");
// }

// interface AppProps{
//   selectdItemFromAdrees:any
// }
// const PDFGenerator:React.FC<AppProps> = ({selectdItemFromAdrees}) => {
//   const [drawerOpen, setDrawerOpen] = useState(false);
//   const [dialogOpen, setDialogOpen] = useState(false);
//   const [errorMessage, setErrorMessage] = useState("");
//   const [PdfBlob, setPdfBlob] = useState(null);
//   const [selectedText, setSelectedText] = useState("");
//   const [loading, setloading] = useState(false);

//   const [textToCopy, setTextToCopy] = useState("");
//   const [emailSettings, setEmailSettings] = useState({
//     email1: "shahzadumarit@gmail.com",
//     email2: "shahzad890.it@outlook.com",
//     email3: "bushraiqbal.it01@gmail.com",
//   });
//   const [onedrivePaths, setOnedrivePaths] = useState({
//     path1: "/Attachments/Demo/Folder1",
//     path2: "/Attachments/Demo/Folder2",
//     path3: "/Attachments/Demo/Folder3",
//   });

//   const [selectedEmails, setSelectedEmails] = useState<string[]>([]); // New state: array of selected emails
//   const [selectedPaths, setSelectedPaths] = useState<string[]>([]);   // New state: array of selected paths
//   const [isSending, setIsSending] = useState(false);
//   const [toast, setToast] = useState({
//     open: false,
//     severity: "info",
//     message: "",
//   });
//   const [emailDetails, setEmailDetails] = useState({
//     body: "",
//     senderEmail: "",
//     sentDate: "",
//     senderInitials: "",
//   });
//   const [showFullText, setShowFullText] = useState(false);
//   const [pdfGenerated, setPdfGenerated] = useState(false); // New state to track PDF generation
// const [failedEmails, setFailedEmails] = useState<string[]>([]); // Store failed email addresses
//   const [failedPaths, setFailedPaths] = useState<string[]>([]); // Store failed OneDrive paths

//   useEffect(() => {
//     // Fetch email details using Office.js
//     Office.context.mailbox.item.body.getAsync("text", (result) => {
//       if (result.status === Office.AsyncResultStatus.Succeeded) {
//         setEmailDetails((prev) => ({
//           ...prev,
//           body: result.value,
//         }));
//       }
//     });

//     const sender = Office.context.mailbox.item.from;
//     setEmailDetails((prev) => ({
//       ...prev,
//       senderEmail: sender.emailAddress,
//       sentDate: Office.context.mailbox.item.dateTimeCreated.toLocaleDateString(),
//       senderInitials: sender.displayName
//         .split(" ")
//         .map((word) => word[0])
//         .join("")
//         .slice(0, 2),
//     }));
//   }, []);
//   useEffect(() => {
//     // Fetch email details using Office.js
//     Office.context.mailbox.item.body.getAsync("text", (result) => {
//       if (result.status === Office.AsyncResultStatus.Succeeded) {
//         setEmailDetails((prev) => ({
//           ...prev,
//           body: result.value,
//         }));
//       }
//     });

//     const sender = Office.context.mailbox.item.from;
//     setEmailDetails((prev) => ({
//       ...prev,
//       senderEmail: sender.emailAddress,
//       sentDate: Office.context.mailbox.item.dateTimeCreated.toLocaleDateString(),
//       senderInitials: sender.displayName
//         .split(" ")
//         .map((word) => word[0])
//         .join("")
//         .slice(0, 2),
//     }));
//   }, [selectdItemFromAdrees]);


//   useEffect(() => {
//     const savedEmailSettings = localStorage.getItem('emailSettings');
//     const savedOnedrivePaths = localStorage.getItem('onedrivePaths');
  
//     if (savedEmailSettings) {
//       setEmailSettings(JSON.parse(savedEmailSettings));
//     }
  
//     if (savedOnedrivePaths) {
//       setOnedrivePaths(JSON.parse(savedOnedrivePaths));
//     }
//   }, []);

  
//   // const handleCopy = (text) => {
//   //   navigator.clipboard.writeText(text).then(() => {
//   //     setTextToCopy(text);
//   //     setDialogOpen(true);
//   //   });
//   // };
//   const handleCopy = (text: string) => {
//     const textArea = document.createElement("textarea");
//     textArea.value = text;
//     document.body.appendChild(textArea);
//     textArea.select();
//     document.execCommand("copy");
//     document.body.removeChild(textArea);
//     setTextToCopy(text);
//     setDialogOpen(true);
//   };
  
//   const handleTextSelect = () => {
//     const selected = window.getSelection()?.toString() || "";
//     if (selected) {
//       setSelectedText(selected);
//       handleCopy(selected);
//     }
//   };

//   const generatePDF = (text, callbackFunction) => {
//     // Get email details from Office context
//     const item = Office.context.mailbox.item;
  
//     // Extract sender's email, subject, and sent date safely
//     const senderEmail = item?.from?.emailAddress || "Unknown Sender";
//     const sentDate = item?.dateTimeCreated
//       ? new Date(item.dateTimeCreated).toLocaleString()
//       : "Unknown Date";
//     const subject = item?.subject || "No Subject";
  
//     // Define PDF structure
//     const docDefinition = {
//       content: [
//         {
//           text: `Subject: ${subject}\nDate: ${sentDate}\nFrom: ${senderEmail}`,
//           style: "header",
//         },
//         { text, style: "body" },
//       ],
//       styles: {
//         header: { fontSize: 14, bold: true, marginBottom: 10 },
//         body: { fontSize: 12, lineHeight: 1.5 },
//       },
//     };
  
//     // Generate PDF and return as blob
//     pdfMake.createPdf(docDefinition).getBlob((blob) => callbackFunction(blob));
//   };
  

//   const handlePDF = (pdfBlob) => {
//     console.log("PDF Blob:", pdfBlob);
//     setPdfBlob(pdfBlob);
//     setPdfGenerated(true); // Set PDF generated state to true
//   };

//   const OpenFullWindow = () => {
//     OpenDialog();
//   };
//   const handleSendEmail = (emails: string[]) => {
//     setloading(true);
//     setFailedEmails([]); // Reset failed emails

//     // Create an array to store promises for each email sending operation
//     const emailPromises: Promise<any>[] = emails.map(email => {
//         const Token = localStorage.getItem("Token");
//         return sendEmailWithAttachments(email, Token)
//             .catch(error => {
//                 console.error(`Failed to send email to ${email}:`, error);
//                 setFailedEmails(prev => [...prev, email]); // Store failed email
//                 // Do not reject here; resolve with the error to keep Promise.all going
//                 return error;
//             });
//     });

//     Promise.all(emailPromises)
//         .then(results => {
//             setloading(false);
//             // Check if all emails were sent successfully
//             const allSuccessful = results.every(result => !(result instanceof Error));

//             if (allSuccessful) {
//                 setToast({
//                     open: true,
//                     severity: "success",
//                     message: "Emails sent successfully!",
//                 });
//             } else {
//                 setToast({
//                     open: true,
//                     severity: "warning", // Use warning severity if some emails failed
//                     message: `Some emails failed to send. See console for details.`,
//                 });
//             }
//         })
//         .catch(error => {
//             console.error("Unexpected error in Promise.all:", error);
//             setloading(false);
//             setToast({
//                 open: true,
//                 severity: "error",
//                 message: "Unexpected error sending emails.",
//             });
//         });
// };


// const LoginAgain=(Item:any,Type)=>{

//   Office.onReady(() => {

//     const dialogOptions: Office.DialogOptions = {
//         height: 40,
//         width: 35,
//         displayInIframe: false
//     }
// console.log(Item);
// // const redirect_uri='https://attachment-sorter.vercel.app/assets/Dialog.html'
// const redirect_uri_For_Local="https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=e5a4342f-c8a5-4185-948d-2e3d485b4822&response_type=token&redirect_uri=https://localhost:3000/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment"
// const redirect_uri_For_LIve="https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=aaa00dc4-7743-467e-8868-596ffff59e05&response_type=token&redirect_uri=https://attachment-sorter.vercel.app/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment";

//     Office.context.ui.displayDialogAsync(
//       redirect_uri_For_Local,
//         dialogOptions,
//         (asyncResult) => {

//             if (asyncResult.status == Office.AsyncResultStatus.Failed) {
//                 console.log("Opening Dialogue Box FAILED Async" + asyncResult.error.message);
//                 return;
//             }

//             const myDialog: Office.Dialog = asyncResult.value;

//             myDialog.addEventHandler(Office.EventType.DialogMessageReceived,
//                 (args: Office.DialogParentMessageReceivedEventArgs) => {
//                     const token = args.message

//                     if (typeof (token) == 'string' && token.length > 50) {
//                         localStorage.setItem("Token", token);
//                         // Re-attempt the failed operations
//             if (Type === "sendEmail") {
//               retryFailedEmails(token);
//           } else if (Type === "OneDrivepath") {
//               retryFailedPaths(token);
//           }

//                         localStorage.setItem('Token', token);
//                     }
//                     else {

//                         //You Might had passed Error parameters. Debug
//                         localStorage.removeItem("Token");
//                         console.error("Got Empty/Incorrect response value @ Dialog window ");
//                         //Debugger here (Check your value & Format. Is length fine?)

//                     }

//                     myDialog.close() //Closes it even incase wrong value

//                     let existingNotifcaiton:any = Office.context.mailbox.item.notificationMessages.getAllAsync(() => { })
//                     //Remove if same tag exist; no spam

//                     Office.context.mailbox.item.notificationMessages.addAsync('Token Notifcation Success ' + existingNotifcaiton.value.length,
//                         {
//                             type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
//                             message: 'Updated Authentication',
//                             icon: "icon_01",
//                             persistent: false

//                         },

//                         () => { })

//                 });

//             myDialog.addEventHandler(Office.EventType.DialogEventReceived, (arg:any) => {

//                 console.error("Dialog Event" + arg.error); //Another Handler if dialog crashses during session time (ie window suddenly closed, access is invalid & etc
//             });

//         }
//     );


// })
// }

// const retryFailedEmails = (token: string) => {
//     setloading(true);

//     Promise.all(failedEmails.map(email =>
//         sendEmailWithAttachments(email, token)
//             .then(() => {
//                 console.log(`Successfully resent email to ${email}`);
//                 setFailedEmails(prev => prev.filter(e => e !== email)); // Remove email from failed list
//             })
//             .catch(error => {
//                 console.error(`Failed to resend email to ${email} after re-authentication:`, error);
//                 // Optionally, handle this failure further (e.g., show an error message)
//             })
//     ))
//     .then(() => {
//         setloading(false);
//         if (failedEmails.length === 0) {
//             setToast({
//                 open: true,
//                 severity: "success",
//                 message: "All emails resent successfully!",
//             });
//         } else {
//             setToast({
//                 open: true,
//                 severity: "warning",
//                 message: `Some emails could not be resent.`,
//             });
//         }
//     });
// };

// const retryFailedPaths = (token: string) => {
//     setloading(true);

//     Promise.all(failedPaths.map(path =>
//         uploadToPath(path, token)
//             .then(() => {
//                 console.log(`Successfully uploaded to path ${path}`);
//                 setFailedPaths(prev => prev.filter(p => p !== path)); // Remove path from failed list
//             })
//             .catch(error => {
//                 console.error(`Failed to upload to path ${path} after re-authentication:`, error);
//                 // Optionally, handle this failure further
//             })
//     ))
//     .then(() => {
//         setloading(false);
//         if (failedPaths.length === 0) {
//             setToast({
//                 open: true,
//                 severity: "success",
//                 message: "All uploads resent successfully!",
//             });
//         } else {
//             setToast({
//                 open: true,
//                 severity: "warning",
//                 message: `Some uploads could not be resent.`,
//             });
//         }
//     });
// };

//   const convertBlobToBase64 = (blob:Blob) => {
//     return new Promise((resolve, reject) => {
//       const reader:any = new FileReader();
//       reader.onloadend = () => {
//         const base64data = reader.result.split(",")[1];
//         resolve(base64data);
//       };
//       reader.onerror = reject;
//       reader.readAsDataURL(blob);
//     });
//   };

//   const sendEmailWithAttachments = async (email: string, Token: string) => {
//     try {
//         setIsSending(true);
//         const base64Pdf = await convertBlobToBase64(PdfBlob);
//         const emailData = {
//             message: {
//                 subject: "Receipt",
//                 body: {
//                     contentType: "Text",
//                     content: selectedEmails.toString(),
//                 },
//                 toRecipients: [
//                     {
//                         emailAddress: {
//                             address: email,
//                         },
//                     },
//                 ],
//                 attachments: [
//                     {
//                         "@odata.type": "#microsoft.graph.fileAttachment",
//                         name: `${Office.context.mailbox.item.subject}.pdf`,
//                         contentType: "application/pdf",
//                         contentBytes: base64Pdf,
//                     },
//                 ],
//             },
//             saveToSentItems: "true",
//         };

//         const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
//             method: "POST",
//             headers: {
//                 "Content-Type": "application/json",
//                 Authorization: `Bearer ${Token}`,
//             },
//             body: JSON.stringify(emailData),
//         });

//         if (response.ok) {
//             console.log(`Email sent successfully to ${email}`);
//             return `Email sent successfully to ${email}`; // Resolve with success message
//         } else {
//             const errorData = await response.json();
//             console.error(`Failed to send email to ${email}:`, errorData);
//             if (errorData.error.code === "InvalidAuthenticationToken") {
//                 LoginAgain(email, 'sendEmail'); // Re-authenticate for the specific email
//             }
//             throw new Error(errorData.error.message || "Failed to send email.");
//         }
//     } catch (error) {
//         console.error(`Error sending email to ${email}:`, error);
//         throw error; // Re-throw the error for handling in the calling function
//     } finally {
//         setIsSending(false);
//     }
// };


// const uploadToPath = async (
//   basePath: string,
//   token: any
 
// ): Promise<string | undefined> => {  // Ensure return type is string | undefined
//   try {
//     const filename = "selectedtext.pdf"; // Consistent filename
//     const fullPath = `${basePath}/${filename}`; // Construct the full path

//     // Define the upload URL with the correct file extension
//     const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:${fullPath}:/content`;

//     // Ensure PdfBlob is a valid Blob object
//     if (!(PdfBlob instanceof Blob)) {
//       console.error("PdfBlob is not a valid Blob object.");
//       // callback(null, "Invalid Blob data");
//       return undefined; // Return undefined if PdfBlob is invalid
//     }

//     // Create a new Blob with the correct MIME type
//     const pdfBlobWithType = new Blob([PdfBlob], { type: "application/pdf" });

//     // Perform the upload using axios
//     const response = await axios.put(uploadUrl, pdfBlobWithType, {
//       headers: {
//         Authorization: `Bearer ${token}`,
//         "Content-Type": "application/pdf",
//       },
//     });

//     // Handle successful upload
//     const successMessage = `PDF uploaded successfully to ${fullPath}`;
//     console.log(successMessage);
//     // callback(successMessage, null);
//      return successMessage
//   } catch (error:any) {
//     // Handle upload error
//      if (error.response && error.response.status === 401) {
//               LoginAgain(basePath, 'OneDrivepath'); // Trigger re-authentication
//           }
//     const errorMessage = `Error uploading to OneDrive path: ${basePath}. Please try again.`;
//     console.error(errorMessage, error);
//     // callback(null, error);
//     return undefined; // Return undefined after error handling
//   }
// };


// const handleUploadPdf=(paths: string[]) => {
//     setloading(true);
//     setFailedPaths([]); // Reset failed paths

//     // Create an array to store promises for each upload operation
//     const uploadPromises: Promise<any>[] = paths.map(path => {
//         const Token = localStorage.getItem("Token");
//         return uploadToPath(path, Token)
//             .catch(error => {
//                 console.error(`Failed to upload to ${path}:`, error);
//                 setFailedPaths(prev => [...prev, path]); // Store failed path
//                 return error;
//             });
//     });

//     Promise.all(uploadPromises)
//         .then(results => {
//             setloading(false);
//             // Check if all uploads were successful
//             const allSuccessful = results.every(result => !(result instanceof Error));

//             if (allSuccessful) {
//                 setToast({
//                     open: true,
//                     severity: "success",
//                     message: "PDFs uploaded successfully to selected paths!",
//                 });
//             } else {
//                 setToast({
//                     open: true,
//                     severity: "warning", // Use warning severity if some uploads failed
//                     message: `Some uploads failed. See console for details.`,
//                 });
//             }
//         })
//         .catch(error => {
//             console.error("Unexpected error in Promise.all:", error);
//             setloading(false);
//             setToast({
//                 open: true,
//                 severity: "error",
//                 message: "Unexpected error uploading PDFs.",
//             });
//         });
// };

// const handleEmailSelection = (event: any) => {
//     setSelectedEmails(event.target.value);
//   };

//   const handlePathSelection = (event: any) => {
//     setSelectedPaths(event.target.value);
//   };
   
//   return (
//     <div>
//     {loading &&(<LoaderApp/>)}
//       <Box>
//         <Typography
//           variant="h5"
//           sx={{
//             marginBottom: 2,
//             textAlign: "center",
//             fontWeight: "bold",
//             display: "flex",
//             alignItems: "center",
//             justifyContent: "space-between",
//             fontSize: "20px",
//           }}
//         >
//           <div>
//             <RiMailSendLine style={{ marginRight: 8, verticalAlign: "middle" }} />
//             Send OR Upload Pdf
//           </div>
//           <IconButton
//             sx={{ float: "right", marginBottom: "5px" }}
//             color="primary"
//             onClick={() => setDrawerOpen(true)}
//           >
//             <RiSettings5Line />
//           </IconButton>
//         </Typography>
        
//         {/* Show this section only if PdfBlob is available */}
//         {pdfGenerated ? (
//           <>
//             {/* OneDrive Upload Section */}
//             <Typography
//           variant="body1"
//           sx={{
//             marginBottom: 2,
//             textAlign: "center",
//             fontWeight: "bold",
//             display: "flex",
//             alignItems: "center",
//             justifyContent: "center",
//             fontSize: "20px",
//           }}
//         >
//           <div style={{fontSize:'medium'}}>
//           Your PDF has been generated successfully! You can now upload it to OneDrive or send it as an email attachment.
//           </div>
//         </Typography>
//         <Divider/>
//             {/* OneDrive Upload Section */}
//             <Box sx={{ marginTop: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'space-between' }}>
//               <Typography variant="h6" sx={{ fontSize: "15px", marginBottom: 1 }}>
//                 Upload to OneDrive
//               </Typography>
//               <FormControl fullWidth sx={{ marginBottom: 2 }}>
//                 <InputLabel id="path-select-label">Select OneDrive Paths</InputLabel>
//                 <Select
//                   labelId="path-select-label"
//                   id="path-select"
//                   multiple
//                   value={selectedPaths}
//                   label="Select OneDrive Paths"
//                   onChange={handlePathSelection}
//                   renderValue={(selected: string[]) => selected.join(', ')}
//                 >
//                   {Object.entries(onedrivePaths).map(([key, path]) => (
//                     <MenuItem key={key} value={path}>
//                       <Checkbox checked={selectedPaths.indexOf(path) > -1} />
//                       <ListItemText primary={path} />
//                     </MenuItem>
//                   ))}
//                 </Select>
//               </FormControl>

//               <Button
//                 variant="contained"
//                 color="primary"
//                 fullWidth
//                 sx={{ fontSize: "10px", marginBottom: 1 }}
//                 disabled={selectedPaths.length === 0}
//                 onClick={() => handleUploadPdf(selectedPaths)}
//               >
//                 Upload to Selected Paths
//               </Button>
//             </Box>
//         <Divider/>

//             {/* Email Sending Section */}
//             <Box sx={{ marginTop: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'space-between' }}>
//               <Typography variant="h6" sx={{ fontSize: "15px", marginBottom: 1 }}>
//                 Send Email
//               </Typography>

//               <FormControl fullWidth sx={{ marginBottom: 2 }}>
//                 <InputLabel id="email-select-label">Select Email Addresses</InputLabel>
//                 <Select
//                   labelId="email-select-label"
//                   id="email-select"
//                   multiple
//                   value={selectedEmails}
//                   label="Select Email Addresses"
//                   onChange={handleEmailSelection}
//                   renderValue={(selected: string[]) => selected.join(', ')}
//                 >
//                   {Object.entries(emailSettings).map(([key, email]) => (
//                     <MenuItem key={key} value={email}>
//                       <Checkbox checked={selectedEmails.indexOf(email) > -1} />
//                       <ListItemText primary={email} />
//                     </MenuItem>
//                   ))}
//                 </Select>
//               </FormControl>

//               <Button
//                 variant="contained"
//                 color="primary"
//                 fullWidth
//                 sx={{ fontSize: "10px", marginBottom: 1 }}
//                 disabled={selectedEmails.length === 0}
//                 onClick={() => handleSendEmail(selectedEmails)}
//               >
//                 Send to Selected Emails
//               </Button>
//             </Box>
//           </>
//         ):(<Card
//           sx={{
//             maxWidth: "100%",
//             boxShadow: "0px 1px 10px 2px #f1f1f1",
//             marginBottom: "10px",
//             maxHeight:'500px',
//             overflow:'auto',
//             scrollbarWidth:'none',
//             scrollBehavior:'smooth'
//           }}
//           onMouseUp={handleTextSelect}
//         >
//           <CardHeader
//             avatar={
//               <Avatar sx={{ bgcolor: red[500] }} aria-label="sender">
//                 {emailDetails.senderInitials}
//               </Avatar>
//             }
//             title={emailDetails.senderEmail}
//             subheader={emailDetails.sentDate}
//           />
//           <CardContent>
//           <Typography variant="body2" sx={{ color: "text.secondary" }}>
//   {emailDetails.body ? (
//     <>
//       {showFullText || emailDetails.body.length <= 200
//         ? emailDetails.body
//         : emailDetails.body.substring(0, 200) + "..."}
      
//       {emailDetails.body.length > 200 && (
//         <Button
//           size="small"
//           color="primary"
//           onClick={() => setShowFullText((prev) => !prev)}
//         >
//           {showFullText ? "Show Less" : "Show More"}
//         </Button>
//       )}
//     </>
//   ) : (
//     "No text available for this email"
//   )}
// </Typography>
//           </CardContent>
//           <CardActions disableSpacing>
//             <IconButton aria-label="copy to clipboard" onClick={() => handleCopy(emailDetails.body)}>
//               <FaRegCopy style={{ fontSize: "19px" }} />
//             </IconButton>
//             <label htmlFor="a">Copy All</label>
  
//           </CardActions>
//         </Card>)}
//       </Box>

//       {/* Drawer for settings */}
//       <Drawer anchor="right" open={drawerOpen} onClose={() => setDrawerOpen(false)}>
//         <Box padding={'10px'}>
//           <Typography variant="h6" sx={{ marginBottom: 2 }}>
//             Edit Email Addresses and OneDrive Paths
//           </Typography>
//           {Object.entries(emailSettings).map(([key, email]) => (
//             <TextField
//             fullWidth
//               key={key}
//               label={`Email ${key.split("email")[1]}`}
//               value={email}
//               onChange={(e) =>
//                 setEmailSettings((prev) => ({ ...prev, [key]: e.target.value }))
//               }
//               sx={{ marginBottom: 2 }}
//             />
//           ))}
//           {Object.entries(onedrivePaths).map(([key, path]) => (
//             <TextField
//               key={key}
//               fullWidth
//               label={`OneDrive Path ${key.split("path")[1]}`}
//               value={path}
//               onChange={(e) =>
//                 setOnedrivePaths((prev) => ({ ...prev, [key]: e.target.value }))
//               }
//               sx={{ marginBottom: 2 }}
//             />
//           ))}
//           <div
//             style={{
//               width: "100%",
//               display: "flex",
//               justifyContent: "space-between",
//               alignItems: "center",
//             }}
//           >
//             <Button variant="outlined" sx={{ width: "49%" }} onClick={() => setDrawerOpen(false)}>
//               Cancel
//             </Button>
//             <Button
//   variant="contained"
//   color="primary"
//   sx={{ width: "49%" }}
//   onClick={() => {
//     localStorage.setItem('emailSettings', JSON.stringify(emailSettings));
//     localStorage.setItem('onedrivePaths', JSON.stringify(onedrivePaths));
//     setDrawerOpen(false);
//   }}
// >
//   Save
// </Button>
//           </div>
//         </Box>
//       </Drawer>

// {errorMessage && (
//         <Alert severity="error" sx={{ marginTop: 2 }} onClose={() => setErrorMessage("")}>
//           {errorMessage}
//         </Alert>
//       )}


//       {/* Modal Dialog */}
//       <Dialog open={dialogOpen} onClose={() => setDialogOpen(false)}>
//         <DialogTitle>Generate PDF</DialogTitle>
//         <DialogContent>
//           <Typography>
//             The text has been copied. Do you want to generate a PDF of the copied text?
//           </Typography>
//         </DialogContent>
//         <DialogActions>
//           <Button onClick={() => setDialogOpen(false)}>No</Button>
//           <Button
//             onClick={() => {
//               generatePDF(textToCopy, handlePDF);
//               setDialogOpen(false);
//             }}
//           >
//             Yes
//           </Button>
//         </DialogActions>
//       </Dialog>
//        <Snackbar
//               open={toast.open}
//               autoHideDuration={50000}
//               onClose={() => setToast({ ...toast, open: false })}
//               sx={{
//                 display: 'flex',
//                 justifyContent: 'center',
//                 alignItems: 'center'
//               }}
//             >
//               <Alert severity={toast.severity as any} onClose={() => setToast({ ...toast, open: false })}>
//                 {toast.message}
//               </Alert>
//             </Snackbar>
//     </div>

//   );
// };

// export default PDFGenerator;













import React, { useState, useEffect } from "react";
import {
  Box,
  Typography,
  TextField,
  Button,
  IconButton,
  Drawer,
  Card,
  CardHeader,
  CardContent,
  Avatar,
  CardActions,
  Alert,
  Snackbar,
  Divider,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  ListItemText,
  Checkbox,
} from "@mui/material";
import { red } from "@mui/material/colors";
import { RiMailSendLine, RiSettings5Line } from "react-icons/ri";
import axios from "axios";
import LoaderApp from "../../Loader/Loader";

interface AppProps {
  selectdItemFromAdrees: any;
}

const EMLHandler: React.FC<AppProps> = ({ selectdItemFromAdrees }) => {
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [emlBlob, setEmlBlob] = useState<Blob | null>(null);
  const [loading, setloading] = useState(false);
  const [emailSettings, setEmailSettings] = useState({
    email1: "test@gmail.com",
    email2: "test@outlook.com",
    email3: "test01@gmail.com",
  });
  const [onedrivePaths, setOnedrivePaths] = useState({
    path1: "/Attachments/Demo/Folder1",
    path2: "/Attachments/Demo/Folder2",
    path3: "/Attachments/Demo/Folder3",
  });
  const [selectedEmails, setSelectedEmails] = useState<string[]>([]);
  const [selectedPaths, setSelectedPaths] = useState<string[]>([]);
  const [toast, setToast] = useState({
    open: false,
    severity: "info",
    message: "",
  });
  const [emailDetails, setEmailDetails] = useState({
    senderEmail: "",
    sentDate: "",
    senderInitials: "",
  });
  const [emlReady, setEmlReady] = useState(false);
  const [failedEmails, setFailedEmails] = useState<string[]>([]);
  const [failedPaths, setFailedPaths] = useState<string[]>([]);

  useEffect(() => {
    const fetchEmailDetails = async () => {
      const sender = Office.context.mailbox.item.from;
      setEmailDetails({
        senderEmail: sender.emailAddress,
        sentDate: Office.context.mailbox.item.dateTimeCreated.toLocaleDateString(),
        senderInitials: sender.displayName
          .split(" ")
          .map((word) => word[0])
          .join("")
          .slice(0, 2),
      });

      try {
        const blob = await getEmailBlob();
        setEmlBlob(blob);
        setEmlReady(true);
      } catch (error) {
        console.error("Error fetching EML:", error);
        setToast({
          open: true,
          severity: "error",
          message: "Failed to retrieve EML file",
        });
      }
    };

    fetchEmailDetails();
  }, [selectdItemFromAdrees]);

  useEffect(() => {
    const savedEmailSettings = localStorage.getItem('emailSettings');
    const savedOnedrivePaths = localStorage.getItem('onedrivePaths');
  
    if (savedEmailSettings) setEmailSettings(JSON.parse(savedEmailSettings));
    if (savedOnedrivePaths) setOnedrivePaths(JSON.parse(savedOnedrivePaths));
  }, []);

  function getItemRestId() {
    if (Office.context.mailbox.diagnostics.hostName === "OutlookIOS") {
      return Office.context.mailbox.item.itemId;
    }
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }

  const getEmailBlob = (): Promise<Blob> => {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(result.error);
          return;
        }

        const token = result.value;
        const getMessageUrl = `${Office.context.mailbox.restUrl}/v2.0/me/messages/${getItemRestId()}/$value`;

        fetch(getMessageUrl, {
          headers: { Authorization: `Bearer ${token}` },
        })
          .then((response) => {
            if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
            return response.blob();
          })
          .then(resolve)
          .catch(reject);
      });
    });
  };

  const convertBlobToBase64 = (blob: Blob): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result as string);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    });
  };

  const sendEmailWithAttachments = async (email: string, Token: string) => {
    try {
      if (!emlBlob) throw new Error("No EML file available");
      
      const subject = Office.context.mailbox.item.subject || "email";
      const base64Eml = await convertBlobToBase64(emlBlob);
      const contentBytes = base64Eml.split(",")[1];

      const emailData = {
        message: {
          subject: `Forwarded: ${subject}`,
          body: {
            contentType: "Text",
            content: "Please find the attached email",
          },
          toRecipients: [{ emailAddress: { address: email } }],
          attachments: [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: `${subject}.eml`,
            contentType: "message/rfc822",
            contentBytes: contentBytes,
          }],
        },
        saveToSentItems: "true",
      };

      const response = await fetch("https://graph.microsoft.com/v1.0/me/sendMail", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${Token}`,
        },
        body: JSON.stringify(emailData),
      });

      if (!response.ok) {
        const errorData = await response.json();
        if (errorData.error.code === "InvalidAuthenticationToken") {
          LoginAgain(email, 'sendEmail');
        }
        throw new Error(errorData.error.message);
      }
      return `EML sent to ${email}`;
    } catch (error) {
      console.error(`Error sending to ${email}:`, error);
      throw error;
    }
  };

  const uploadToOneDrive = async (path: string, token: string) => {
    if (!emlBlob) {
      setToast({ open: true, severity: "error", message: "No EML file available" });
      return;
    }

    try {
      let sanitizedPath = path.trim().replace(/^\/+|\/+$/g, ""); // Remove extra slashes
      if (!sanitizedPath) throw new Error("Invalid OneDrive path");

      const subject = Office.context.mailbox.item.subject || "email";
      const filename = encodeURIComponent(`${subject.replace(/\//g, "-")}.eml`); // Replace invalid characters
      const fullPath = `${sanitizedPath}/${filename}`;

      // Ensure the folder exists before uploading
      await createFolderIfNotExists(sanitizedPath, token);

      const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fullPath)}:/content`;
      await axios.put(uploadUrl, emlBlob, {
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "message/rfc822" },
      });

      setToast({ open: true, severity: "success", message: `EML uploaded successfully!` });
    } catch (error) {
      setToast({ open: true, severity: "error", message: "Upload failed. Check path and token." });
    }
  };

  const createFolderIfNotExists = async (path: string, token: string) => {
    try {
      const folderUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(path)}:/`;
      await axios.get(folderUrl, { headers: { Authorization: `Bearer ${token}` } });
    } catch (error) {
      if (error.response?.status === 404) {
        await axios.post(
          `https://graph.microsoft.com/v1.0/me/drive/root/children`,
          { name: path, folder: {} },
          { headers: { Authorization: `Bearer ${token}` } }
        );
      } else {
        throw error;
      }
    }
  };

  const handleSendEmail = (emails: string[]) => {
    setloading(true);
    setFailedEmails([]);

    Promise.all(emails.map(email => {
      const Token = localStorage.getItem("Token");
      return sendEmailWithAttachments(email, Token || "")
        .catch(error => {
          setFailedEmails(prev => [...prev, email]);
          return error;
        });
    })).then(results => {
      setloading(false);
      const allSuccessful = results.every(result => !(result instanceof Error));
      setToast({
        open: true,
        severity: allSuccessful ? "success" : "warning",
        message: allSuccessful 
          ? "Emails sent successfully!" 
          : "Some emails failed to send",
      });
    }).catch(__error => {
      setloading(false);
      setToast({
        open: true,
        severity: "error",
        message: "Unexpected error sending emails",
      });
    });
  };

  const handleUpload = (paths: string[]) => {
    setloading(true);
    setFailedPaths([]);

    Promise.all(paths.map(path => {
      const Token = localStorage.getItem("Token");
      return uploadToOneDrive(path, Token || "")
        .catch(error => {
          setFailedPaths(prev => [...prev, path]);
          return error;
        });
    })).then(results => {
      setloading(false);
      const allSuccessful = results.every(result => !(result instanceof Error));
      setToast({
        open: true,
        severity: allSuccessful ? "success" : "warning",
        message: allSuccessful 
          ? "EMLs uploaded successfully!" 
          : "Some uploads failed",
      });
    }).catch(__error => {
      setloading(false);
      setToast({
        open: true,
        severity: "error",
        message: "Unexpected error uploading",
      });
    });
  };

  const LoginAgain = (__item: any, type: string) => {
    Office.onReady(() => {
      const dialogOptions: Office.DialogOptions = {
        height: 40,
        width: 35,
        displayInIframe: false
      };
      // const redirect_uri = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=e5a4342f-c8a5-4185-948d-2e3d485b4822&response_type=token&redirect_uri=https://localhost:3000/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment";
      const clientId = process.env.NODE_ENV === 'development' 
      ? 'e5a4342f-c8a5-4185-948d-2e3d485b4822' 
      : 'ce3ab054-f9ae-4f8d-b530-c7dd621e7213';
    
    const redirectURI = process.env.NODE_ENV === 'development' 
      ? 'https://localhost:3000/assets/Dialog.html' 
      : 'https://attachment-sorter.vercel.app/assets/Dialog.html';
    
    const authURL = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${redirectURI}&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment`;
    
      Office.context.ui.displayDialogAsync(authURL, dialogOptions, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) return;
        const myDialog: Office.Dialog = asyncResult.value;
        myDialog.addEventHandler(Office.EventType.DialogMessageReceived, (args:any) => {
          const token = args.message;
          if (typeof token === 'string' && token.length > 50) {
            localStorage.setItem("Token", token);
            if (type === "sendEmail") retryFailedEmails(token);
            if (type === "OneDrivepath") retryFailedPaths(token);
          }
          myDialog.close();
        });
      });
    });
  };

  const retryFailedEmails = (token: string) => {
      setloading(true);
  
      Promise.all(failedEmails.map(email =>
          sendEmailWithAttachments(email, token)
              .then(() => {
                  console.log(`Successfully resent email to ${email}`);
                  setFailedEmails(prev => prev.filter(e => e !== email)); // Remove email from failed list
              })
              .catch(error => {
                  console.error(`Failed to resend email to ${email} after re-authentication:`, error);
                  // Optionally, handle this failure further (e.g., show an error message)
              })
      ))
      .then(() => {
          setloading(false);
          if (failedEmails.length === 0) {
              setToast({
                  open: true,
                  severity: "success",
                  message: "All emails resent successfully!",
              });
          } else {
              setToast({
                  open: true,
                  severity: "warning",
                  message: `Some emails could not be resent.`,
              });
          }
      });
  };
  
  const retryFailedPaths = (token: string) => {
      setloading(true);
  
      Promise.all(failedPaths.map(path =>
        uploadToOneDrive(path, token)
              .then(() => {
                  console.log(`Successfully uploaded to path ${path}`);
                  setFailedPaths(prev => prev.filter(p => p !== path)); // Remove path from failed list
              })
              .catch(error => {
                  console.error(`Failed to upload to path ${path} after re-authentication:`, error);
                  // Optionally, handle this failure further
              })
      ))
      .then(() => {
          setloading(false);
          if (failedPaths.length === 0) {
              setToast({
                  open: true,
                  severity: "success",
                  message: "All uploads resent successfully!",
              });
          } else {
              setToast({
                  open: true,
                  severity: "warning",
                  message: `Some uploads could not be resent.`,
              });
          }
      });
  };



  return (
    <div>
      {loading && <LoaderApp />}
      <Box>
        <Typography variant="h5" sx={{
          marginBottom: 2,
          textAlign: "center",
          fontWeight: "bold",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          fontSize: "20px",
        }}>
          <div>
            <RiMailSendLine style={{ marginRight: 8 }} />
            EML File Handler
          </div>
          <IconButton onClick={() => setDrawerOpen(true)}>
            <RiSettings5Line />
          </IconButton>
        </Typography>

        {emlReady ? (
          <>
            <Typography variant="body1" sx={{
              marginBottom: 2,
              textAlign: "center",
              fontWeight: "bold",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              fontSize: "20px",
            }}>
              <div style={{ fontSize: 'medium' }}>
                EML file ready! Choose an action below:
              </div>
            </Typography>
            <Divider />

            <Box sx={{ marginTop: 2 }}>
              <FormControl fullWidth sx={{ marginBottom: 2 }}>
                <InputLabel>OneDrive Paths</InputLabel>
                <Select
                  multiple
                  value={selectedPaths}
                  onChange={(e) => setSelectedPaths(e.target.value as string[])}
                  renderValue={(selected) => (selected as string[]).join(', ')}
                >
                  {Object.entries(onedrivePaths).map(([key, path]) => (
                    <MenuItem key={key} value={path}>
                      <Checkbox checked={selectedPaths.includes(path)} />
                      <ListItemText primary={path} />
                    </MenuItem>
                  ))}
                </Select>
              </FormControl>
              <Button
                variant="contained"
                fullWidth
                onClick={() => handleUpload(selectedPaths)}
                disabled={!selectedPaths.length}
              >
                Upload to OneDrive
              </Button>
            </Box>

            <Divider sx={{ marginY: 2 }} />

            <Box sx={{ marginTop: 2 }}>
              <FormControl fullWidth sx={{ marginBottom: 2 }}>
                <InputLabel>Email Addresses</InputLabel>
                <Select
                  multiple
                  value={selectedEmails}
                  onChange={(e) => setSelectedEmails(e.target.value as string[])}
                  renderValue={(selected) => (selected as string[]).join(', ')}
                >
                  {Object.entries(emailSettings).map(([key, email]) => (
                    <MenuItem key={key} value={email}>
                      <Checkbox checked={selectedEmails.includes(email)} />
                      <ListItemText primary={email} />
                    </MenuItem>
                  ))}
                </Select>
              </FormControl>
              <Button
                variant="contained"
                fullWidth
                onClick={() => handleSendEmail(selectedEmails)}
                disabled={!selectedEmails.length}
              >
                Send as Email Attachment
              </Button>
            </Box>
          </>
        ) : (
          <Card sx={{
            maxWidth: "100%",
            boxShadow: "0px 1px 10px 2px #f1f1f1",
            marginBottom: "10px",
          }}>
            <CardHeader
              avatar={<Avatar sx={{ bgcolor: red[500] }}>{emailDetails.senderInitials}</Avatar>}
              title={emailDetails.senderEmail}
              subheader={emailDetails.sentDate}
            />
            <CardContent>
              <Typography variant="body2" sx={{ color: "text.secondary" }}>
                Loading EML file...
              </Typography>
            </CardContent>
          </Card>
        )}

        <Drawer anchor="right" open={drawerOpen} onClose={() => setDrawerOpen(false)}>
          <Box padding={'10px'}>
            <Typography variant="h6" sx={{ marginBottom: 2 }}>
              Configuration Settings
            </Typography>
            {Object.entries(emailSettings).map(([key, email]) => (
              <TextField
                fullWidth
                key={key}
                label={`Email ${key.slice(-1)}`}
                value={email}
                onChange={(e) => setEmailSettings(prev => ({ ...prev, [key]: e.target.value }))}
                sx={{ marginBottom: 2 }}
              />
            ))}
            {Object.entries(onedrivePaths).map(([key, path]) => (
              <TextField
                key={key}
                fullWidth
                label={`Path ${key.slice(-1)}`}
                value={path}
                onChange={(e) => setOnedrivePaths(prev => ({ ...prev, [key]: e.target.value }))}
                sx={{ marginBottom: 2 }}
              />
            ))}
            <Box sx={{ display: 'flex', gap: 2 }}>
              <Button fullWidth variant="outlined" onClick={() => setDrawerOpen(false)}>
                Cancel
              </Button>
              <Button fullWidth variant="contained" onClick={() => {
                localStorage.setItem('emailSettings', JSON.stringify(emailSettings));
                localStorage.setItem('onedrivePaths', JSON.stringify(onedrivePaths));
                setDrawerOpen(false);
              }}>
                Save
              </Button>
            </Box>
          </Box>
        </Drawer>

        <Snackbar
          open={toast.open}
          autoHideDuration={6000}
          onClose={() => setToast(prev => ({ ...prev, open: false }))}
        >
          <Alert severity={toast.severity as any}>
            {toast.message}
          </Alert>
        </Snackbar>
      </Box>
    </div>
  );
};

export default EMLHandler;





// import React, { useState, useEffect } from "react";
// import {
//   Box,
//   Typography,
//   TextField,
//   Button,
//   IconButton,
//   Drawer,
//   Select,
//   MenuItem,
//   FormControl,
//   InputLabel,
//   ListItemText,
//   Checkbox,
//   Snackbar,
//   Alert,
// } from "@mui/material";
// import { RiMailSendLine, RiSettings5Line } from "react-icons/ri";
// import axios from "axios";
// import LoaderApp from "../../Loader/Loader";

// interface AppProps {
//   selectdItemFromAdrees: any;
// }

// const EMLHandler: React.FC<AppProps> = ({ selectdItemFromAdrees }) => {
//   const [drawerOpen, setDrawerOpen] = useState(false);
//   const [emlBlob, setEmlBlob] = useState<Blob | null>(null);
//   const [loading, setLoading] = useState(false);
//   const [toast, setToast] = useState({ open: false, severity: "info", message: "" });

//   // Email recipients & OneDrive paths
//   const [onedrivePaths, setOnedrivePaths] = useState({
//     path1: "/Attachments/Demo/Folder1",
//     path2: "/Attachments/Demo/Folder2",
//     path3: "/Attachments/Demo/Folder3",
//   });

//   const [selectedPaths, setSelectedPaths] = useState<string[]>([]);

//   useEffect(() => {
//     const savedPaths = localStorage.getItem("onedrivePaths");
//     if (savedPaths) setOnedrivePaths(JSON.parse(savedPaths));
//     fetchEmailBlob();
//   }, [selectdItemFromAdrees]);

//   const fetchEmailBlob = async () => {
//     setLoading(true);
//     try {
//       const blob = await getEmailBlob();
//       setEmlBlob(blob);
//     } catch (error) {
//       setToast({ open: true, severity: "error", message: "Failed to retrieve EML file" });
//     }
//     setLoading(false);
//   };

//   const getEmailBlob = (): Promise<Blob> => {
//     return new Promise((resolve, reject) => {
//       Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
//         if (result.status !== Office.AsyncResultStatus.Succeeded) {
//           reject(result.error);
//           return;
//         }

//         const token = result.value;
//         const getMessageUrl = `${Office.context.mailbox.restUrl}/v2.0/me/messages/${Office.context.mailbox.convertToRestId(
//           Office.context.mailbox.item.itemId,
//           Office.MailboxEnums.RestVersion.v2_0
//         )}/$value`;

//         fetch(getMessageUrl, {
//           headers: { Authorization: `Bearer ${token}` },
//         })
//           .then((response) => response.blob())
//           .then(resolve)
//           .catch(reject);
//       });
//     });
//   };

//   const uploadToOneDrive = async (path: string, token: string) => {
//     if (!emlBlob) {
//       setToast({ open: true, severity: "error", message: "No EML file available" });
//       return;
//     }

//     try {
//       let sanitizedPath = path.trim().replace(/^\/+|\/+$/g, ""); // Remove extra slashes
//       if (!sanitizedPath) throw new Error("Invalid OneDrive path");

//       const subject = Office.context.mailbox.item.subject || "email";
//       const filename = encodeURIComponent(`${subject.replace(/\//g, "-")}.eml`); // Replace invalid characters
//       const fullPath = `${sanitizedPath}/${filename}`;

//       // Ensure the folder exists before uploading
//       await createFolderIfNotExists(sanitizedPath, token);

//       const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(fullPath)}:/content`;
//       await axios.put(uploadUrl, emlBlob, {
//         headers: { Authorization: `Bearer ${token}`, "Content-Type": "message/rfc822" },
//       });

//       setToast({ open: true, severity: "success", message: `EML uploaded successfully!` });
//     } catch (error) {
//       setToast({ open: true, severity: "error", message: "Upload failed. Check path and token." });
//     }
//   };

//   const createFolderIfNotExists = async (path: string, token: string) => {
//     try {
//       const folderUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(path)}:/`;
//       await axios.get(folderUrl, { headers: { Authorization: `Bearer ${token}` } });
//     } catch (error) {
//       if (error.response?.status === 404) {
//         await axios.post(
//           `https://graph.microsoft.com/v1.0/me/drive/root/children`,
//           { name: path, folder: {} },
//           { headers: { Authorization: `Bearer ${token}` } }
//         );
//       } else {
//         throw error;
//       }
//     }
//   };

//   return (
//     <div>
//       {loading && <LoaderApp />}
//       <Box>
//         <Typography
//           variant="h5"
//           sx={{
//             marginBottom: 2,
//             textAlign: "center",
//             fontWeight: "bold",
//             display: "flex",
//             alignItems: "center",
//             justifyContent: "space-between",
//             fontSize: "20px",
//           }}
//         >
//           <div>
//             <RiMailSendLine style={{ marginRight: 8 }} />
//             EML File Handler
//           </div>
//           <IconButton onClick={() => setDrawerOpen(true)}>
//             <RiSettings5Line />
//           </IconButton>
//         </Typography>

//         {/* OneDrive Path Selection */}
//         <FormControl fullWidth sx={{ marginBottom: 2 }}>
//           <InputLabel>OneDrive Paths</InputLabel>
//           <Select
//             multiple
//             value={selectedPaths}
//             onChange={(e) => setSelectedPaths(e.target.value as string[])}
//             renderValue={(selected) => (selected as string[]).join(", ")}
//           >
//             {Object.entries(onedrivePaths).map(([key, path]) => (
//               <MenuItem key={key} value={path}>
//                 <Checkbox checked={selectedPaths.includes(path)} />
//                 <ListItemText primary={path} />
//               </MenuItem>
//             ))}
//           </Select>
//         </FormControl>

//         <Button variant="contained" fullWidth onClick={() => uploadToOneDrive(selectedPaths[0], localStorage.getItem("Token") || "")}>
//           Upload to OneDrive
//         </Button>

//         {/* Settings Drawer */}
//         <Drawer anchor="right" open={drawerOpen} onClose={() => setDrawerOpen(false)}>
//           <Box padding={"10px"}>
//             <Typography variant="h6">Configuration Settings</Typography>
//             {Object.entries(onedrivePaths).map(([key, path]) => (
//               <TextField
//                 key={key}
//                 fullWidth
//                 label={`Path ${key.slice(-1)}`}
//                 value={path}
//                 onChange={(e) => setOnedrivePaths((prev) => ({ ...prev, [key]: e.target.value })) }
//                 sx={{ marginBottom: 2 }}
//               />
//             ))}
//             <Button fullWidth variant="contained" onClick={() => localStorage.setItem("onedrivePaths", JSON.stringify(onedrivePaths))}>
//               Save
//             </Button>
//           </Box>
//         </Drawer>

//         {/* Snackbar Notifications */}
//         <Snackbar open={toast.open} autoHideDuration={6000} onClose={() => setToast({ ...toast, open: false })}>
//           <Alert severity={toast.severity as any}>{toast.message}</Alert>
//         </Snackbar>
//       </Box>
//     </div>
//   );
// };

// export default EMLHandler;
