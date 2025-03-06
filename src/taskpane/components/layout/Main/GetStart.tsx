import React, { useEffect } from 'react';
import { useNavigate } from 'react-router-dom';
import { Button, Box, Typography, Paper, Container } from '@mui/material';
import { useState } from 'react';
const GetStart:React.FC= () => {
   const [AccessToken, setAccessToken] = useState('');
  useEffect(() => {
    let storeToken=localStorage.getItem('Token')
if(storeToken){
  navigate('home')
}
   }, []);
  const navigate = useNavigate();

  // const GoToNextPage = (__event) => {

  
  //  //local
  //   var a ="https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=e5a4342f-c8a5-4185-948d-2e3d485b4822&response_type=token&redirect_uri=https://localhost:3000/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment";
  //  //live
  //   // var a = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=aaa00dc4-7743-467e-8868-596ffff59e05&response_type=token&redirect_uri=https://attachment-sorter.vercel.app/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment";
  //   //client azure
  //   // var a = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=8a59fa27-967f-4bcf-ab82-e89c82fa0fa5&response_type=token&redirect_uri=https://attachment-sorter.vercel.app/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment";

  //   Office.context.ui.displayDialogAsync(a, { height: 80, width: 60 }, function (asyncResult) {
  //   let  Logindialog = asyncResult.value;
  //     Logindialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg:any) {
  //     let  token = arg.message;
  //     const currentTime = Date.now(); // Current time in ms
  //     const tokenLifetime = 5 * 60 * 1000; // Token valid for 5 minutes (in ms)

  //     // Save token and time in sessionStorage
  //     localStorage.setItem('Token', token);
  //     navigate('Home');
  //       Logindialog.close();
  //     });
  //   });
  // };



const GoToNextPage=()=>{

  Office.onReady(() => {

    const dialogOptions: Office.DialogOptions = {
        height: 60,
        width: 60,
        displayInIframe: false
    }
    // const redirect_uri_For_Local="https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=e5a4342f-c8a5-4185-948d-2e3d485b4822&response_type=token&redirect_uri=https://localhost:3000/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment"
    // const redirect_uri_For_LIve="https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=aaa00dc4-7743-467e-8868-596ffff59e05&response_type=token&redirect_uri=https://attachment-sorter.vercel.app/assets/Dialog.html&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment";
 
    const clientId = process.env.NODE_ENV === 'development' 
    ? 'e5a4342f-c8a5-4185-948d-2e3d485b4822' 
    : '567985e0-287c-432c-8755-90fcd55789f6';
  
  const redirectURI = process.env.NODE_ENV === 'development' 
    ? 'https://localhost:3000/assets/Dialog.html' 
    : 'https://attachment-sorter.vercel.app/assets/Dialog.html';
  
  const authURL = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=${clientId}&response_type=token&redirect_uri=${redirectURI}&scope=mail.send+Files.ReadWrite.All+openid+profile+email&response_mode=fragment`;
  



    Office.context.ui.displayDialogAsync(
      authURL,      dialogOptions,
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
navigate('home')
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
    <Box
      sx={{
        display: 'flex',
        flexDirection: 'column',
        justifyContent: 'center',
        alignItems: 'center',
        minHeight: '100vh',
      }}
    >
      <Container maxWidth="sm">
        <Box
          sx={{
            textAlign: 'center',
            backgroundColor: 'white',
          }}
        >
          {/* Logo Section */}
          <Box sx={{ marginBottom: 3 }}>
            <img
              width={200}
              height={'auto'}
              src={require('../../../../../assets/Logo.png')} // Replace with your logo
              alt="Attachment Sorter Logo"
            />
          </Box>

          {/* Add-in Title */}
          <Typography
            variant="h5"
            component="h1"
            sx={{
              fontFamily: 'Roboto, sans-serif',
              fontWeight: 700,
              color: '#333',
              marginBottom: 2,
            }}
          >
            Attachment Sorter
          </Typography>

          {/* Description */}
          <Typography
            variant="body1"
            sx={{
              fontFamily: 'Roboto, sans-serif',
              color: '#555',
              lineHeight: 1.6,
              fontSize: '16px',
              marginBottom: 3,
              textAlign: 'center',
            }}
          >
            Simplify your email attachment management with Attachment Sorter.  
            Effortlessly organize and send attachments to your desired email or OneDrive folder, directly from Outlook.  
            Streamline your workflow with customized settings to manage attachments efficiently.
          </Typography>

          {/* Features */}
         

          {/* Get Started Button */}
          <Button
            variant="contained"
            onClick={GoToNextPage}
            sx={{
              color: 'white',
              padding: '9px 25px',
              fontSize: '16px',
              fontWeight: 600,
              width: '100%',
              borderRadius: 3,
              boxShadow: 2,
              
            }}
          >
            Get Started
          </Button>
        </Box>
      </Container>
    </Box>
  );
};

export default GetStart;
