import React, { useEffect, useState } from "react";
import { Box, Typography, Button, TextField, Avatar, Card, CardActions, CardContent, CardHeader, IconButton } from "@mui/material";
import { red } from "@mui/material/colors";
import { FaRegCopy } from "react-icons/fa";
import { FaArrowsLeftRight } from "react-icons/fa6";
import { RiMailSendLine, RiSettings5Line } from "react-icons/ri";

const AppDialog = () => {
  const [selectedText, setSelectedText] = useState("");
    const [textToCopy, setTextToCopy] = useState("");
  
const [emailDetails, setEmailDetails] = useState({
    body: "",
    senderEmail: "",
    sentDate: "",
    senderInitials: "",
  });
  const [showFullText, setShowFullText] = useState(false);  // State to toggle show more/less
const [dialogOpen, setDialogOpen] = useState(false);
  useEffect(() => {
    // Fetch email details using Office.js
    Office.context.mailbox.item.body.getAsync("text", (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        setEmailDetails((prev) => ({
          ...prev,
          body: result.value,
        }));
      }
    });

    const sender = Office.context.mailbox.item.from;
    setEmailDetails((prev) => ({
      ...prev,
      senderEmail: sender.emailAddress,
      sentDate: Office.context.mailbox.item.dateTimeCreated.toLocaleDateString(),
      senderInitials: sender.displayName
        .split(" ")
        .map((word) => word[0])
        .join("")
        .slice(0, 2),
    }));
  }, []);

  const handleCopy = (text) => {
    navigator.clipboard.writeText(text).then(() => {
      setTextToCopy(text);
      setDialogOpen(true);
    });
  };
  // Dummy data for email details

  // Handle sending selected text to the parent
  const sendToParent = () => {
    Office.context.ui.messageParent(JSON.stringify({ selectedText }));
  };
  const handleTextSelect = () => {
    const selected = window.getSelection()?.toString() || "";
    if (selected) {
      setSelectedText(selected);
      handleCopy(selected);
    }
  };
  return (
    <Box>
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
        <RiMailSendLine style={{ marginRight: 8, verticalAlign: "middle" }} />
        Send OR Upload Pdf
      </div>
    </Typography>

    <Card
      sx={{
        maxWidth: "100%",
        boxShadow: "0px 1px 10px 2px #f1f1f1",
        marginBottom: "10px",
        maxHeight:'500px',
        overflow:'auto',
        scrollbarWidth:'none',
        scrollBehavior:'smooth'
      }}
      onMouseUp={handleTextSelect}
    >
      <CardHeader
        avatar={
          <Avatar sx={{ bgcolor: red[500] }} aria-label="sender">
            {emailDetails.senderInitials}
          </Avatar>
        }
        title={emailDetails.senderEmail}
        subheader={emailDetails.sentDate}
      />
      <CardContent>
        <Typography variant="body2" sx={{ color: "text.secondary" }}>
          {showFullText
            ? emailDetails.body
            : emailDetails.body.substring(0, 200) + "..."}
          <Button
            size="small"
            color="primary"
            onClick={() => setShowFullText((prev) => !prev)}  // Toggle show more/less
          >
            {showFullText ? "Show Less" : "Show More"}
          </Button>
        </Typography>
      </CardContent>
      <CardActions disableSpacing>
        <IconButton aria-label="copy to clipboard" onClick={() => handleCopy(emailDetails.body)}>
          <FaRegCopy style={{ fontSize: "19px" }} />
        </IconButton>
      </CardActions>
    </Card>
  </Box>
  );
};

export default AppDialog;
