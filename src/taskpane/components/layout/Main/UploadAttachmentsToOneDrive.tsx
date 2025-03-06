// import React, { useEffect, useState } from "react";
// import {
//   Box,
//   Typography,
//   Button,
//   List,
//   ListItem,
//   ListItemIcon,
//   ListItemText,
//   Snackbar,
//   Alert,
//   ListItemButton,
//   Checkbox,
//   Divider,
//   Drawer,
//   IconButton,
//   TextField,
// } from "@mui/material";
// import { FaCloudUploadAlt } from "react-icons/fa";
// import axios from "axios";
// import LoaderApp from "../../Loader/Loader";
// import { RiSettings5Line } from "react-icons/ri";
// import {
//   FaDatabase,
//   FaFile,
//   FaFileAlt,
//   FaFileArchive,
//   FaFileAudio,
//   FaFileCode,
//   FaFileCsv,
//   FaFileExcel,
//   FaFileImage,
//   FaFilePdf,
//   FaFilePowerpoint,
//   FaFileVideo,
//   FaFileWord,
// } from "react-icons/fa";

// interface AppProps {
//   selectdItemFromAdrees: any;
// }

// interface Attachment {
//   id: string;
//   name: string;
//   contentType: string;
//   contentBytes?: string;
// }

// const UploadAttachmentsToOneDrive: React.FC<AppProps> = ({ selectdItemFromAdrees }) => {
//   const [attachments, setAttachments] = useState<Attachment[]>([]);
//   const [selectedAttachments, setSelectedAttachments] = useState<string[]>([]);
//   const [selectedPaths, setSelectedPaths] = useState<string[]>([]);
//   const [isUploading, setIsUploading] = useState(false);
//   const [uploadSuccess, setUploadSuccess] = useState(false);
//   const [errorMessage, setErrorMessage] = useState("");
//   const [loading, setLoading] = useState(false);
//   const [drawerOpen, setDrawerOpen] = useState(false);

//   // OneDrive paths state
//   const [uploadPaths, setUploadPaths] = useState({
//     path1: "Attachments",
//     path2: "Path2",
//     path3: "Path3",
//   });

//   // Load attachments and paths from local storage
//   useEffect(() => {
//     const fetchAttachments = async () => {
//       setLoading(true);
//       const emailAttachments = Office.context.mailbox.item.attachments || [];
//       const preparedAttachments:any = await prepareAttachments(emailAttachments);
//       setAttachments(preparedAttachments);
//       setLoading(false);
//     };

//     const savedPaths = localStorage.getItem('uploadPaths');
//     if (savedPaths) {
//       setUploadPaths(JSON.parse(savedPaths));
//     }

//     fetchAttachments();
//   }, [selectdItemFromAdrees]);

//   // Fetch attachment content
//   const fetchAttachmentContent = async (attachmentId: string): Promise<Blob> => {
//     return new Promise((resolve, reject) => {
//       Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
//         if (result.status === Office.AsyncResultStatus.Succeeded) {
//           const base64Content = result.value.content;
//           const binaryString = atob(base64Content);
//           const byteArray = new Uint8Array(binaryString.length);
//           for (let i = 0; i < binaryString.length; i++) {
//             byteArray[i] = binaryString.charCodeAt(i);
//           }
//           const blob = new Blob([byteArray], { type: result.value.format });
//           resolve(blob);
//         } else {
//           console.error("Failed to fetch attachment content:", result.error.message);
//           reject(result.error.message);
//         }
//       });
//     });
//   };

//   // Prepare attachments
//   const prepareAttachments = async (attachments: Office.AttachmentDetails[]) => {
//     return Promise.all(
//       attachments.map(async (attachment) => {
//         try {
//           const blob = await fetchAttachmentContent(attachment.id);
//           return {
//             id: attachment.id,
//             name: attachment.name,
//             contentType: attachment.contentType,
//             contentBytes: blob,
//           };
//         } catch (error) {
//           console.error("Error preparing attachment:", attachment.name, error);
//           return { id: attachment.id, name: attachment.name, contentType: attachment.contentType };
//         }
//       })
//     );
//   };

//   // Handle attachments checkbox
//   const toggleAttachmentSelection = (id: string) => {
//     setSelectedAttachments((prev) =>
//       prev.includes(id) ? prev.filter((attachmentId) => attachmentId !== id) : [...prev, id]
//     );
//   };

//   // Handle paths checkbox
//   const togglePathSelection = (path: string) => {
//     setSelectedPaths((prev) =>
//       prev.includes(path) ? prev.filter((p) => p !== path) : [...prev, path]
//     );
//   };

  // // Upload to OneDrive
  // const uploadToPath = async (path: string, Token: string, callback: (data: any, error: any) => void) => {
  //   if (selectedAttachments.length === 0) {
  //     setErrorMessage("No attachments selected for upload.");
  //     return;
  //   }
  //   setIsUploading(true);
  //   setErrorMessage("");
  //   try {
  //     for (const attachment of attachments.filter((a) => selectedAttachments.includes(a.id))) {
  //       if (!attachment.contentBytes) continue;

  //       // Prepare the upload URL
  //       const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${path}/${attachment.name}:/content`;

  //       // Upload the Blob directly
  //       await axios.put(uploadUrl, attachment.contentBytes, {
  //         headers: {
  //           Authorization: `Bearer ${Token}`,
  //           "Content-Type": attachment.contentType,
  //         },
  //       });

  //       console.log(`${attachment.name} uploaded successfully to ${path}`);
  //     }
  //     callback(`${attachments.length} files uploaded successfully to ${path}`, null);
  //     setUploadSuccess(true);
  //     setSelectedAttachments([]);
  //   } catch (error) {
  //     setErrorMessage("Error uploading to OneDrive. Please try again.");
  //     console.error("Upload error:", error);
  //     callback(null, error);
  //   } finally {
  //     setIsUploading(false);
  //   }
  // };

//   // Handle drawer save
//   const handleSave = () => {
//     localStorage.setItem('uploadPaths', JSON.stringify(uploadPaths));
//     setDrawerOpen(false);
//   };

//   // Handle drawer cancel
//   const handleCancel = () => {
//     setDrawerOpen(false);
//   };

//   // Get file icon
//   const getIcon = (type: string) => {
//     switch (type) {
//       case "image/png":
//       case "image/jpeg":
//       case "image/jpg":
//       case "image/gif":
//       case "image/webp":
//       case "image/svg+xml":
//       case "image/bmp":
//         return <FaFileImage color="blue" fontSize={"xx-large"} />;
//       case "application/pdf":
//         return <FaFilePdf color="red" fontSize={"xx-large"} />;
//       case "application/msword":
//       case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
//         return <FaFileWord color="blue" fontSize={"xx-large"} />;
//       case "application/vnd.ms-excel":
//       case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
//       case "text/csv":
//         return <FaFileExcel color="green" fontSize={"xx-large"} />;
//       case "text/csv":
//         return <FaFileCsv color="teal" fontSize={"xx-large"} />;
//       case "application/vnd.ms-powerpoint":
//       case "application/vnd.openxmlformats-officedocument.presentationml.presentation":
//         return <FaFilePowerpoint color="orange" fontSize={"xx-large"} />;
//       case "application/zip":
//       case "application/x-7z-compressed":
//       case "application/x-rar-compressed":
//       case "application/x-tar":
//       case "application/gzip":
//         return <FaFileArchive color="brown" fontSize={"xx-large"} />;
//       case "text/plain":
//         return <FaFileAlt color="gray" fontSize={"xx-large"} />;
//       case "audio/mpeg":
//       case "audio/wav":
//       case "audio/ogg":
//       case "audio/aac":
//       case "audio/x-midi":
//         return <FaFileAudio color="purple" fontSize={"xx-large"} />;
//       case "video/mp4":
//       case "video/x-msvideo":
//       case "video/mpeg":
//       case "video/ogg":
//       case "video/webm":
//       case "video/x-matroska":
//         return <FaFileVideo color="red" fontSize={"xx-large"} />;
//       case "text/html":
//       case "application/javascript":
//       case "application/json":
//       case "text/css":
//       case "application/xml":
//         return <FaFileCode color="green" fontSize={"xx-large"} />;
//       case "application/vnd.ms-access":
//       case "application/x-sqlite3":
//       case "application/x-msdownload":
//         return <FaDatabase color="brown" fontSize={"xx-large"} />;
//       default:
//         return <FaFile color="black" fontSize={"xx-large"} />;
//     }
//   };

//   // Handle upload
//   const handleUploadAttachments = () => {
//     const Token = localStorage.getItem("Token");
//     selectedPaths.forEach((path) => {
//       uploadToPath(path, Token, (data, error) => {
//         if (data) {
//           setUploadSuccess(true);
//           setSelectedAttachments([]);
//         }
//         if (error) {
//           if (error.status === 401) {
//             setErrorMessage("Token is Expire Please Login.");
//             LoginAgain(path);
//           } else {
//             setErrorMessage(error.message);
//           }
//         }
//       });
//     });
//   };

//   // Handle login
//   const LoginAgain = (path: string) => {
//     const authUrl = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=e5a4342f-c8a5-4185-948d-2e3d485b4822&response_type=token&redirect_uri=https://localhost:3000/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment";

//     Office.context.ui.displayDialogAsync(authUrl, { height: 80, width: 60 }, (asyncResult) => {
//       const loginDialog = asyncResult.value;
//       loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg: any) => {
//         const token = arg.message;
//         localStorage.setItem('Token', token);
//         uploadToPath(path, token, (data, error) => {
//           if (data) {
//             setUploadSuccess(true);
//             setSelectedAttachments([]);
//           }
//           if (error && error.status === 401) {
//             setErrorMessage("Token is Expire Please Login.");
//             LoginAgain(path);
//           }
//         });
//         loginDialog.close();
//       });
//     });
//   };

//   return (
//     <Box sx={{ maxWidth: 600, margin: "auto", backgroundColor: "#fff" }}>
//       {loading && <LoaderApp />}
//       <Typography
//         variant="h5"
//         sx={{
//           marginBottom: 2,
//           textAlign: "center",
//           fontWeight: "bold",
//           display: "flex",
//           alignItems: "center",
//           justifyContent: "space-between",
//           fontSize: "20px",
//         }}
//       >
//         <div>
//           <FaCloudUploadAlt style={{ marginRight: 8, verticalAlign: "middle" }} />
//           Upload To OneDrive
//         </div>
//         <IconButton sx={{ float: "right", marginBottom: "5px" }} color="primary" onClick={() => setDrawerOpen(true)}>
//           <RiSettings5Line />
//         </IconButton>
//       </Typography>

//       <Box sx={{ maxHeight: 300, overflow: "auto" }}>
//         <List>
//           {attachments.length === 0 && !loading && (
//             <Box textAlign="center" p={2}>
//               No Attachment Found with the selected mail
//             </Box>
//           )}
//           {attachments.map((attachment) => (
//             <React.Fragment key={attachment.id}>
//               <ListItem disablePadding>
//                 <Checkbox
//                   checked={selectedAttachments.includes(attachment.id)}
//                   onChange={() => toggleAttachmentSelection(attachment.id)}
//                 />
//                 <ListItemButton>
//                   <ListItemIcon>{getIcon(attachment.contentType)}</ListItemIcon>
//                   <ListItemText sx={{ fontSize: "14px" }} primary={attachment.name} />
//                 </ListItemButton>
//               </ListItem>
//               <Divider />
//             </React.Fragment>
//           ))}
//         </List>
//       </Box>

//       {/* Paths List */}
//       <Typography variant="h6" sx={{ marginTop: 2, fontSize: "15px" }}>
//         Select Upload Paths
//       </Typography>
//       <List>
//         {Object.entries(uploadPaths).map(([key, path]) => (
//           <ListItem key={key} disablePadding>
//             <Checkbox checked={selectedPaths.includes(path)} onChange={() => togglePathSelection(path)} />
//             <ListItemText primary={path} />
//           </ListItem>
//         ))}
//       </List>

//       {/* Upload Button */}
//       <Button
//         variant="contained"
//         color="primary"
//         fullWidth
//         sx={{
//           "&.Mui-disabled": {
//             backgroundColor: "#93C0FE",
//             color: "white",
//           },
//           marginTop: 2,
//         }}
//         disabled={isUploading || !selectedAttachments.length || !selectedPaths.length}
//         onClick={handleUploadAttachments}
//         startIcon={<FaCloudUploadAlt />}
//       >
//         Upload to Selected Paths
//       </Button>

//       {/* Success and Error Messages */}
//       {errorMessage && (
//         <Alert severity="error" sx={{ marginTop: 2 }} onClose={() => setErrorMessage("")}>
//           {errorMessage}
//         </Alert>
//       )}

//       <Snackbar
//         open={uploadSuccess}
//         autoHideDuration={3000}
//         onClose={() => setUploadSuccess(false)}
//         anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
//       >
//         <Alert severity="success" variant="filled">
//           Attachments uploaded successfully!
//         </Alert>
//       </Snackbar>

//       {/* Settings Drawer */}
//       <Drawer anchor="right" open={drawerOpen} onClose={() => setDrawerOpen(false)}>
//         <Box sx={{ padding: 2 }}>
//           <Typography variant="h6">Edit Upload Paths</Typography>

//           <TextField
//             label="Path 1"
//             fullWidth
//             margin="normal"
//             value={uploadPaths.path1}
//             onChange={(e) => setUploadPaths({ ...uploadPaths, path1: e.target.value })}
//           />

//           <TextField
//             label="Path 2"
//             fullWidth
//             margin="normal"
//             value={uploadPaths.path2}
//             onChange={(e) => setUploadPaths({ ...uploadPaths, path2: e.target.value })}
//           />

//           <TextField
//             label="Path 3"
//             fullWidth
//             margin="normal"
//             value={uploadPaths.path3}
//             onChange={(e) => setUploadPaths({ ...uploadPaths, path3: e.target.value })}
//           />

//           <div style={{ width: "100%", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
//             <Button variant="outlined" sx={{ width: "49%" }} onClick={handleCancel}>
//               Cancel
//             </Button>
//             <Button variant="contained" color="primary" sx={{ width: "49%" }} onClick={handleSave}>
//               Save
//             </Button>
//           </div>
//         </Box>
//       </Drawer>
//     </Box>
//   );
// };

// export default UploadAttachmentsToOneDrive;












import React, { useEffect, useState } from "react";
import {
  Box,
  Typography,
  Button,
  List,
  ListItem,
  ListItemIcon,
  ListItemText,
  Snackbar,
  Alert,
  ListItemButton,
  Checkbox,
  Divider,
  Drawer,
  IconButton,
  TextField,
} from "@mui/material";
import { FaCloudUploadAlt } from "react-icons/fa";
import axios from "axios";
import LoaderApp from "../../Loader/Loader";
import { RiSettings5Line } from "react-icons/ri";
import {
  FaDatabase,
  FaFile,
  FaFileAlt,
  FaFileArchive,
  FaFileAudio,
  FaFileCode,
  FaFileCsv,
  FaFileExcel,
  FaFileImage,
  FaFilePdf,
  FaFilePowerpoint,
  FaFileVideo,
  FaFileWord,
} from "react-icons/fa";

interface AppProps {
  selectdItemFromAdrees: any;
}

interface Attachment {
  id: string;
  name: string;
  contentType: string;
  contentBytes?: string;
}

const UploadAttachmentsToOneDrive: React.FC<AppProps> = ({ selectdItemFromAdrees }) => {
  const [attachments, setAttachments] = useState<Attachment[]>([]);
  const [selectedAttachments, setSelectedAttachments] = useState<string[]>([]);
  const [selectedPaths, setSelectedPaths] = useState<string[]>([]);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadSuccess, setUploadSuccess] = useState(false);
  const [errorMessage, setErrorMessage] = useState("");
  const [loading, setLoading] = useState(false);
  const [drawerOpen, setDrawerOpen] = useState(false);

  // OneDrive paths state
  const [uploadPaths, setUploadPaths] = useState({
    path1: "/Attachments/Demo/Folder1",
    path2: "/Attachments/Demo/Folder2",
    path3: "/Attachments/Demo/Folder3",
  });

  // Load attachments and paths from local storage
  useEffect(() => {
    const fetchAttachments = async () => {
      setLoading(true);
      const emailAttachments = Office.context.mailbox.item.attachments || [];
      const preparedAttachments:any = await prepareAttachments(emailAttachments);
      setAttachments(preparedAttachments);
      setLoading(false);
    };

    const savedPaths = localStorage.getItem('uploadPaths');
    if (savedPaths) {
      setUploadPaths(JSON.parse(savedPaths));
    }

    fetchAttachments();
  }, [selectdItemFromAdrees]);

  // Fetch attachment content
  const fetchAttachmentContent = async (attachmentId: string): Promise<Blob> => {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.getAttachmentContentAsync(attachmentId, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const base64Content = result.value.content;
          const binaryString = atob(base64Content);
          const byteArray = new Uint8Array(binaryString.length);
          for (let i = 0; i < binaryString.length; i++) {
            byteArray[i] = binaryString.charCodeAt(i);
          }
          const blob = new Blob([byteArray], { type: result.value.format });
          resolve(blob);
        } else {
          console.error("Failed to fetch attachment content:", result.error.message);
          reject(result.error.message);
        }
      });
    });
  };

  // Prepare attachments
  const prepareAttachments = async (attachments: Office.AttachmentDetails[]) => {
    return Promise.all(
      attachments.map(async (attachment) => {
        try {
          const blob = await fetchAttachmentContent(attachment.id);
          return {
            id: attachment.id,
            name: attachment.name,
            contentType: attachment.contentType,
            contentBytes: blob,
          };
        } catch (error) {
          console.error("Error preparing attachment:", attachment.name, error);
          return { id: attachment.id, name: attachment.name, contentType: attachment.contentType };
        }
      })
    );
  };

  // Handle attachments checkbox
  const toggleAttachmentSelection = (id: string) => {
    setSelectedAttachments((prev) =>
      prev.includes(id) ? prev.filter((attachmentId) => attachmentId !== id) : [...prev, id]
    );
  };

  // Handle paths checkbox
  const togglePathSelection = (path: string) => {
    setSelectedPaths((prev) =>
      prev.includes(path) ? prev.filter((p) => p !== path) : [...prev, path]
    );
  };

  // Upload to OneDrive
  // const uploadToPath = async (path: string, Token: string, attachment: Attachment, callback: (data: any, error: any) => void) => {
  //   setIsUploading(true);
  //   setErrorMessage("");
  //   try {
  //     if (!attachment.contentBytes) {
  //       callback(null, "Attachment content is missing.");
  //       return;
  //     }

  //     const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${path}/${attachment.name}:/content`;

  //     await axios.put(uploadUrl, attachment.contentBytes, {
  //       headers: {
  //         Authorization: `Bearer ${Token}`,
  //         "Content-Type": attachment.contentType,
  //       },
  //     });
  //     callback(`${attachment.name} uploaded successfully to ${path}`, null);
  //   } catch (error) {
  //     setErrorMessage(`Error uploading ${attachment.name} to ${path}. Please try again.`);
  //     console.error(`Upload error for ${attachment.name} to ${path}:`, error);
  //     callback(null, error);
  //   } finally {
  //     setIsUploading(false);
  //   }
  // };






  const uploadToPath = async (path, token, callback) => {
    if (selectedAttachments.length === 0) {
      setErrorMessage("No attachments selected for upload.");
      return;
    }
  
    setIsUploading(true);
    setErrorMessage("");
  
    try {
      // 1. Sanitize and split the base path
      const sanitizedPath = path.replace(/^\/+|\/+$/g, ""); // Remove leading/trailing slashes
      const pathSegments = sanitizedPath.split("/").filter(Boolean);
  
      // 2. Create the folder structure (if it doesn't exist)
      // We'll build the current path iteratively.
      let currentPath = "";
      for (const segment of pathSegments) {
        currentPath = currentPath ? `${currentPath}/${segment}` : segment;
        // Build the encoded path (each segment encoded)
        const encodedPath = currentPath
          .split("/")
          .map((s) => encodeURIComponent(s))
          .join("/");
        // Use the trailing colon format for the GET request
        const folderUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedPath}:`;
  
        try {
          // Check if folder exists
          await axios.get(folderUrl, {
            headers: { Authorization: `Bearer ${token}` },
          });
        } catch (error) {
          if (error.response?.status === 404) {
            // Folder doesn't exist: create it.
            // Determine the parent folder URL. If currentPath has no slash, the parent is root.
            const parentPath = currentPath.lastIndexOf("/") === -1 ? "" : currentPath.substring(0, currentPath.lastIndexOf("/"));
            let parentUrl;
            if (parentPath === "") {
              parentUrl = `https://graph.microsoft.com/v1.0/me/drive/root/children`;
            } else {
              const encodedParentPath = parentPath
                .split("/")
                .map((s) => encodeURIComponent(s))
                .join("/");
              parentUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedParentPath}:/children`;
            }
  
            await axios.post(
              parentUrl,
              {
                name: segment,
                folder: {},
                "@microsoft.graph.conflictBehavior": "rename",
              },
              { headers: { Authorization: `Bearer ${token}` } }
            );
          } else {
            throw error;
          }
        }
      }
  
      // 3. Upload each selected attachment to the created path
      // Build the final encoded folder path for uploading
      const encodedUploadPath = sanitizedPath
        .split("/")
        .map((s) => encodeURIComponent(s))
        .join("/");
  
      // Loop through each selected attachment
      for (const attachment of attachments.filter((a) => selectedAttachments.includes(a.id))) {
        if (!attachment.contentBytes) continue;
  
        const encodedFileName = encodeURIComponent(attachment.name);
        // const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodedUploadPath}/${encodedFileName}:/content`;
        const uploadUrl = `https://graph.microsoft.com/v1.0/me/drive/root:${sanitizedPath}/${encodedFileName}:/content`;

        await axios.put(uploadUrl, attachment.contentBytes, {
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": attachment.contentType,
          },
        });
      }
  
      callback(`${selectedAttachments.length} files uploaded successfully to ${sanitizedPath}`, null);
      setUploadSuccess(true);
      setSelectedAttachments([]);
    } catch (error) {
      setErrorMessage("Error uploading to OneDrive. Please try again.");
      console.error("Upload error:", error);
      callback(null, error);
    } finally {
      setIsUploading(false);
    }
  };
  
  
  






  const uploadAllAttachmentsToPaths = async (Token: string) => {
    setIsUploading(true);
    setErrorMessage("");

    try {
      const uploadPromises = selectedPaths.flatMap(path =>
        attachments
          .filter(a => selectedAttachments.includes(a.id))
          .map(_attachment =>
            new Promise((_resolve, _reject) => {
              uploadToPath(path, Token, (data, error) => {
                if (data) {
                  setUploadSuccess(true);
                  setSelectedAttachments([]);
                }
                if (error && error.status === 401) {
                  setErrorMessage("Token is Expire Please Login.");
                  LoginAgain(path);
                }
              });
            })
          )
      );

      await Promise.all(uploadPromises);

      setUploadSuccess(true);
      setSelectedAttachments([]);
      console.log("All attachments uploaded successfully to selected paths");
    } catch (error) {
      setErrorMessage("Error uploading to OneDrive. Please check the console for details.");
      console.error("Upload error:", error);
    } finally {
      setIsUploading(false);
    }
  };

  const handleUploadAttachments = () => {
    const Token = localStorage.getItem("Token");
    uploadAllAttachmentsToPaths(Token);
  };

  // Handle drawer save
  const handleSave = () => {
    localStorage.setItem('uploadPaths', JSON.stringify(uploadPaths));
    setDrawerOpen(false);
  };

  // Handle drawer cancel
  const handleCancel = () => {
    setDrawerOpen(false);
  };

  // Get file icon
  const getIcon = (type: string) => {
    switch (type) {
      case "image/png":
      case "image/jpeg":
      case "image/jpg":
      case "image/gif":
      case "image/webp":
      case "image/svg+xml":
      case "image/bmp":
        return <FaFileImage color="blue" fontSize={"xx-large"} />;
      case "application/pdf":
        return <FaFilePdf color="red" fontSize={"xx-large"} />;
      case "application/msword":
      case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        return <FaFileWord color="blue" fontSize={"xx-large"} />;
      case "application/vnd.ms-excel":
      case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
      case "text/csv":
        return <FaFileExcel color="green" fontSize={"xx-large"} />;
      case "text/csv":
        return <FaFileCsv color="teal" fontSize={"xx-large"} />;
      case "application/vnd.ms-powerpoint":
      case "application/vnd.openxmlformats-officedocument.presentationml.presentation":
        return <FaFilePowerpoint color="orange" fontSize={"xx-large"} />;
      case "application/zip":
      case "application/x-7z-compressed":
      case "application/x-rar-compressed":
      case "application/x-tar":
      case "application/gzip":
        return <FaFileArchive color="brown" fontSize={"xx-large"} />;
      case "text/plain":
        return <FaFileAlt color="gray" fontSize={"xx-large"} />;
      case "audio/mpeg":
      case "audio/wav":
      case "audio/ogg":
      case "audio/aac":
      case "audio/x-midi":
        return <FaFileAudio color="purple" fontSize={"xx-large"} />;
      case "video/mp4":
      case "video/x-msvideo":
      case "video/mpeg":
      case "video/ogg":
      case "video/webm":
      case "video/x-matroska":
        return <FaFileVideo color="red" fontSize={"xx-large"} />;
      case "text/html":
      case "application/javascript":
      case "application/json":
      case "text/css":
      case "application/xml":
        return <FaFileCode color="green" fontSize={"xx-large"} />;
      case "application/vnd.ms-access":
      case "application/x-sqlite3":
      case "application/x-msdownload":
        return <FaDatabase color="brown" fontSize={"xx-large"} />;
      default:
        return <FaFile color="black" fontSize={"xx-large"} />;
    }
  };

  // Handle login
  const LoginAgain=(path:any)=>{

    Office.onReady(() => {
  
      const dialogOptions: Office.DialogOptions = {
          height: 40,
          width: 35,
          displayInIframe: false
      }
      const redirect_uri_For_Local="https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=e5a4342f-c8a5-4185-948d-2e3d485b4822&response_type=token&redirect_uri=https://localhost:3000/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment"
      const redirect_uri_For_LIve="https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=aaa00dc4-7743-467e-8868-596ffff59e05&response_type=token&redirect_uri=https://shahzadumar-w.github.io/OutlookAddin_AttachmentSorter/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment";
  
      Office.context.ui.displayDialogAsync(
        redirect_uri_For_LIve, dialogOptions,
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
                          uploadToPath(path, token, (data, error) => {
                                      if (data) {
                                        setUploadSuccess(true);
                                        setSelectedAttachments([]);
                                      }
                                      if (error && error.status === 401) {
                                        setErrorMessage("Token is Expire Please Login.");
                                        LoginAgain(path);
                                      }
                                    });
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
  return (
    <Box sx={{ maxWidth: 600, margin: "auto", backgroundColor: "#fff" }}>
      {loading && <LoaderApp />}
      <Typography
        variant="h5"
        sx={{
          marginBottom: 2,
          textAlign: "center",
          fontWeight: "bold",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          fontSize: "20px",
        }}
      >
        <div>
          <FaCloudUploadAlt style={{ marginRight: 8, verticalAlign: "middle" }} />
          Upload To OneDrive
        </div>
        <IconButton sx={{ float: "right", marginBottom: "5px" }} color="primary" onClick={() => setDrawerOpen(true)}>
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
                  <ListItemText sx={{ fontSize: "14px" }} primary={attachment.name} />
                </ListItemButton>
              </ListItem>
              <Divider />
            </React.Fragment>
          ))}
        </List>
      </Box>

      {/* Paths List */}
      <Typography variant="h6" sx={{ marginTop: 2, fontSize: "15px" }}>
        Select Upload Paths
      </Typography>
      <List>
        {Object.entries(uploadPaths).map(([key, path]) => (
          <ListItem key={key} disablePadding>
            <Checkbox checked={selectedPaths.includes(path)} onChange={() => togglePathSelection(path)} />
            <ListItemText primary={path} />
          </ListItem>
        ))}
      </List>

      {/* Upload Button */}
      <Button
        variant="contained"
        color="primary"
        fullWidth
        sx={{
          "&.Mui-disabled": {
            backgroundColor: "#93C0FE",
            color: "white",
          },
          marginTop: 2,
        }}
        disabled={isUploading || !selectedAttachments.length || !selectedPaths.length}
        onClick={handleUploadAttachments}
        startIcon={<FaCloudUploadAlt />}
      >
        Upload to Selected Paths
      </Button>

      {/* Success and Error Messages */}
      {errorMessage && (
        <Alert severity="error" sx={{ marginTop: 2 }} onClose={() => setErrorMessage("")}>
          {errorMessage}
        </Alert>
      )}

      <Snackbar
        open={uploadSuccess}
        autoHideDuration={3000}
        onClose={() => setUploadSuccess(false)}
        anchorOrigin={{ vertical: "bottom", horizontal: "center" }}
      >
        <Alert severity="success" variant="filled">
          Attachments uploaded successfully!
        </Alert>
      </Snackbar>

      {/* Settings Drawer */}
      <Drawer anchor="right" open={drawerOpen} onClose={() => setDrawerOpen(false)}>
        <Box sx={{ padding: 2 }}>
          <Typography variant="h6">Edit Upload Paths</Typography>

          <TextField
            label="Path 1"
            fullWidth
            margin="normal"
            value={uploadPaths.path1}
            onChange={(e) => setUploadPaths({ ...uploadPaths, path1: e.target.value })}
          />

          <TextField
            label="Path 2"
            fullWidth
            margin="normal"
            value={uploadPaths.path2}
            onChange={(e) => setUploadPaths({ ...uploadPaths, path2: e.target.value })}
          />

          <TextField
            label="Path 3"
            fullWidth
            margin="normal"
            value={uploadPaths.path3}
            onChange={(e) => setUploadPaths({ ...uploadPaths, path3: e.target.value })}
          />

          <div style={{ width: "100%", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <Button variant="outlined" sx={{ width: "49%" }} onClick={handleCancel}>
              Cancel
            </Button>
            <Button variant="contained" color="primary" sx={{ width: "49%" }} onClick={handleSave}>
              Save
            </Button>
          </div>
        </Box>
      </Drawer>
    </Box>
  );
};

export default UploadAttachmentsToOneDrive;