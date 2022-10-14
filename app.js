import soundfile from './TeamsRingTone.mp3'; 
  
const { PublicClientApplication } = require('@azure/msal-browser');

const { CallClient, VideoStreamRenderer, LocalVideoStream } = require('@azure/communication-calling');
const { AzureCommunicationTokenCredential } = require('@azure/communication-common');
 
const { AzureLogger, setLogLevel } = require("@azure/logger");

import { Features} from "@azure/communication-calling";
import { apply } from 'file-loader';

// Set the log level and output
setLogLevel('verbose');
AzureLogger.log = (...args) => {
    console.log(...args);
};

const publicClientApplication = new PublicClientApplication({
	auth: {
        clientId: "57cddb84-8833-48b2-b8b7-baddba6db02f",
        authority: "https://login.microsoftonline.com/b9d2f243-b20e-44ca-aaba-b179b9963fe2",
    },
    system: {
        tokenRenewalOffsetSeconds: 900 // 15 minutes (by default 5 minutes)
    }
});

// UI objects
let transferTargetPhone = document.getElementById('transfer-target-phone');
let transferCallButton = document.getElementById('transfer-call-button');
let acceptCallButton = document.getElementById('accept-call-button');
let hangUpCallButton = document.getElementById('hangup-call-button');

document.getElementById("transfer-call-button-image").src="hstgray.jpg";
transferCallButton.disabled = true;
transferTargetPhone.disabled = true;	

var actionMsg = {};
var infoMsg = {};
var messageMsg = {};

actionMsg = {type: "action", command: "idle", datetime: getFormattedDate()};
sendPostMessage(actionMsg);

actionMsg = {type: "action", command: "transferoff", datetime: getFormattedDate()};
sendPostMessage(actionMsg);

// app objects
let ACSToken;
let ACSTokenExpires;
let TeamsUserName;
let TeamsUserId;
let TeamsToken;
let TeamsUserEmail;
let tokenCredential;
let callAgent;
let deviceManager;

let call1;
let call2;
let callTransferApi1;
let callTransferApi2;

let incomingCall;
let incomingCallId;
let callerInfo;
let ringTone;
let btnstatus = "00";

ringTone = new Audio(soundfile);
ringTone.loop = true;

function getFormattedDate() {
    var d = new Date();

    d = d.getFullYear() + "/" + ('0' + (d.getMonth() + 1)).slice(-2) + "/" + ('0' + d.getDate()).slice(-2) + " " + ('0' + d.getHours()).slice(-2) + ":" + ('0' + d.getMinutes()).slice(-2) + ":" + ('0' + d.getSeconds()).slice(-2);

    return d;
}

function sendPostMessage(pmsg) {

	try {
		var infoMsg = pmsg;
		var childWindow = document.getElementById("selvbetjeningfr");
		childWindow = childWindow ? childWindow.contentWindow : null;
		childWindow.postMessage(infoMsg, "*");
		console.log(infoMsg);
	}
	catch (error) {
	   console.error(error);
	}

}

const fetchTokenFromMyServerForUser = async function (username) {

	console.log("[REFRESHTOKEN] fetchTokenFromMyServerForUser arrived.");

    // Refresh the Azure AD access token of the Teams User
    let teamsTokenResponse = await refreshAadToken(username);

	if (teamsTokenResponse != null) {
		console.log("[REFRESHTOKEN] Accessing getTokenForTeamsUser");

		var url = "https://waframegettoken.azurewebsites.net/getTokenForTeamsUser";
		//var url = "https://wagettoken.azurewebsites.net/getTokenForTeamsUser";
	
		// Exchange the Azure AD access token of the Teams User for a Communication Identity access token
		const response = await fetch(url,
			{
				method: "POST",
				body: JSON.stringify({ teamsToken: teamsTokenResponse.accessToken }),
				headers: { 'Content-Type': 'application/json' }
			});
	
		console.log("[REFRESHTOKEN] fetch finished.");
	
		if (response.ok) {
	
			console.log("[REFRESHTOKEN] response ok: " + data.communicationIdentityToken);
	
			const data = await response.json();
			return data.communicationIdentityToken;
		}
	}
}

const refreshAadToken = async function (username) {

	console.log("[REFRESHTOKEN] refreshAadToken arrived for " + username);

    //let account = (await publicClientApplication.getTokenCache().getAllAccounts()).find(u => u.username === username);

	await publicClientApplication.handleRedirectPromise();
	var account1 = publicClientApplication.getAccountByUsername(TeamsUserEmail);
	console.log("[REFRESHTOKEN] account1 found: " + account1);

	var account = publicClientApplication.getAllAccounts().find(u => u.username === username);
	console.log("[REFRESHTOKEN] account found: " + account);

    let tokenResponse = null;

	if (account == undefined) {
		console.log("[REFRESHTOKEN] account still not found.");
		return tokenResponse;
	}

    const renewRequest = {
        scopes: ["https://auth.msft.communication.azure.com/Teams.ManageCalls"],
        account: account,
        forceRefresh: true // Force-refresh the token
    };

	console.log("[REFRESHTOKEN] getting token silent.");

    // Try to get the token silently without the user's interaction    
    await publicClientApplication.acquireTokenSilent(renewRequest).then(renewResponse => {
        tokenResponse = renewResponse;
		console.log("[REFRESHTOKEN] new token response " + tokenResponse.accessToken);
    }).catch(async (error) => {
		console.log("[REFRESHTOKEN] " + error);
    });

    return tokenResponse;
}

// Create an instance of CallClient. Initialize a CallAgent instance with a CommunicationUserCredential via created CallClient. 
async function CreateInstance() {
    try {

        const callClient = new CallClient(); 
        
		// short-lived
		//tokenCredential = new AzureCommunicationTokenCredential(ACSToken);

		tokenCredential = new AzureCommunicationTokenCredential({
            tokenRefresher: async () => fetchTokenFromMyServerForUser(TeamsUserEmail),
			refreshProactively: true,
            token: ACSToken
        });
		
        callAgent = await callClient.createCallAgent(tokenCredential);
		
		document.getElementById("accept-call-button-image").src="hsgray.jpg";
		document.getElementById("hangup-call-button-image").src="hshugray.jpg";
		btnstatus = "00";

		actionMsg = {type: "action", command: "idle", datetime: getFormattedDate()};
		sendPostMessage(actionMsg);

		document.getElementById("transfer-call-button-image").src="hstgray.jpg";
		transferCallButton.disabled = true;
		transferTargetPhone.disabled = true;

		document.getElementById("connectedLabel").innerHTML = "Environment is initiated!";
		infoMsg = {type: "info", message: "Environment is initiated!", datetime: getFormattedDate()};
		sendPostMessage(infoMsg);

		actionMsg = {type: "action", command: "transferoff", datetime: getFormattedDate()};
		sendPostMessage(actionMsg);
		        
		// Set up a audio device to use.
        deviceManager = await callClient.getDeviceManager();
        await deviceManager.askDevicePermission({ audio: true });
		
        // Listen for an incoming call to accept.
        callAgent.on('incomingCall', async (args) => {
            try {
					incomingCall = args.incomingCall;	

					incomingCall.on('callEnded', args => {
						console.log(args.callEndReason);

						document.getElementById("accept-call-button-image").src="hsgray.jpg";
						document.getElementById("hangup-call-button-image").src="hshugray.jpg";
						btnstatus = "00";

						actionMsg = {type: "action", command: "idle", datetime: getFormattedDate()};
						sendPostMessage(actionMsg);
						
						document.getElementById("transfer-call-button-image").src="hstgray.jpg";
						transferCallButton.disabled = true;
						transferTargetPhone.disabled = true;	

						actionMsg = {type: "action", command: "transferoff", datetime: getFormattedDate()};
						sendPostMessage(actionMsg);
						
						document.getElementById("connectedLabel").innerHTML = "Current call ended. Environment is initiated.";
						infoMsg = {type: "info", message: "Current call ended. Environment is initiated.", datetime: getFormattedDate()};
						sendPostMessage(infoMsg);

						infoMsg = {type: "message", status: "HangUp", callid: incomingCall.id, datetime: getFormattedDate()};
						sendPostMessage(infoMsg);

						// stop ring tune
						ringTone.pause();
					});
					
					// Get incoming call ID
					incomingCallId = incomingCall.id
        
					var callKind = incomingCall.callerInfo.identifier.kind;
					var callPhone = incomingCall.callerInfo.identifier.phoneNumber;
					var callDisplayName = incomingCall.callerInfo.displayName;
					
					var callToShow = callDisplayName;
					if (callPhone != callDisplayName) 
					{
						if (callPhone != undefined) {
							callToShow += ", " + callPhone;
						}
					}
	
					// Get information about caller
					callerInfo = callToShow;// + " (" + callKind + ")";
	
					document.getElementById("accept-call-button-image").src="hsgr.jpg";
					document.getElementById("hangup-call-button-image").src="hshugray.jpg";
					btnstatus = "10";
					
					actionMsg = {type: "action", command: "callringing", datetime: getFormattedDate()};
					sendPostMessage(actionMsg);
					
					document.getElementById("connectedLabel").innerHTML = "Ring... Ring... Incoming call from: " + callerInfo + ", Call Id: " + incomingCallId;
					infoMsg = {type: "info", message: "Ring... Ring... Incoming call from: " + callerInfo + ", Call Id: " + incomingCallId, datetime: getFormattedDate()};
					sendPostMessage(infoMsg);

					// play ring tune
					ringTone.load();

					try {
						ringTone.play();
					}
					catch (error) {
						console.log(error);
					}

					messageMsg = {type: "message", status: "Presented", callid: incomingCall.id, datetime: getFormattedDate()};
					sendPostMessage(messageMsg);

            } catch (error) {
                console.error(error);
            }
        });
        //initializeCallAgentButton.disabled = true;
    } catch(error) {
        console.error(error);
    }
}

const queryString = window.location.search;
if ((queryString == null) || (queryString == "")) 
{
	document.getElementById("connectedLabel").innerHTML = "Acquiring token, please wait...";
	infoMsg = {type: "info", message: "Acquiring token, please wait...", datetime: getFormattedDate()};
	sendPostMessage(infoMsg);

	window.location.href = "https://waframegettoken.azurewebsites.net/";
	//window.location.href = "https://wagettoken.azurewebsites.net/";
}

const urlParams = new URLSearchParams(queryString);
if ((urlParams == null) || (urlParams == "")) 
{
	document.getElementById("connectedLabel").innerHTML = "Acquiring token, please wait...";
	infoMsg = {type: "info", message: "Acquiring token, please wait...", datetime: getFormattedDate()};
	sendPostMessage(infoMsg);

	window.location.href = "https://waframegettoken.azurewebsites.net/";
	//window.location.href = "https://wagettoken.azurewebsites.net/";
}
else 
{
	document.getElementById("connectedLabel").innerHTML = "Initiating environment, please wait...";
	infoMsg = {type: "info", message: "Initiating environment, please wait...", datetime: getFormattedDate()};
	sendPostMessage(infoMsg);

	ACSToken = urlParams.get('token');
	ACSTokenExpires = urlParams.get('expires');

	actionMsg = {type: "action", command: "tokenexpirestime", tokenexpires: ACSTokenExpires, datetime: getFormattedDate()};
	sendPostMessage(actionMsg);

	TeamsUserEmail = urlParams.get('email');
	TeamsUserId = urlParams.get('userid');
	TeamsUserName = urlParams.get('name');
	TeamsToken = urlParams.get('teamstoken');

    CreateInstance();
}

async function goAcceptTheCall() {
	if (btnstatus != "10") {
		return;
	}

    try {
        call1 = await incomingCall.accept();

		// const threadId = call.info.threadId;

		var callKind = incomingCall.callerInfo.identifier.kind;
		var callPhone = incomingCall.callerInfo.identifier.phoneNumber;
		var callDisplayName = incomingCall.callerInfo.displayName;
		
		var callToShow = callDisplayName;
		if (callPhone != callDisplayName) 
		{
			if (callPhone != undefined) {
				callToShow += ", " + callPhone;
			}
		}

		callerInfo = callToShow;// + " (" + callKind + ")";

		callTransferApi1 = call1.feature(Features.Transfer);
		callTransferApi1.on('transferRequested', args => {
			console.log(`Receive transfer request: ${args.targetParticipant}`);
			args.accept();
		});

		// stop ring tune
		ringTone.pause();
		
		document.getElementById("accept-call-button-image").src="hsgray.jpg";
		document.getElementById("hangup-call-button-image").src="hshured.jpg";
		btnstatus = "01";

		actionMsg = {type: "action", command: "callactive", datetime: getFormattedDate()};
		sendPostMessage(actionMsg);

		document.getElementById("transfer-call-button-image").src="hstgr.jpg";
		transferCallButton.disabled = false;
		transferTargetPhone.disabled = false;

		actionMsg = {type: "action", command: "transferon", datetime: getFormattedDate()};
		sendPostMessage(actionMsg);

		var vThreadIdInfo = "";
		if (call1.info.threadId != undefined) {
			vThreadIdInfo =  ", ThreadId found: " + call1.info.threadId;
		}
		
		document.getElementById("connectedLabel").innerHTML = "Connected to: " + callerInfo + ", Call Id: " + incomingCallId + vThreadIdInfo;
		infoMsg = {type: "info", message: "Connected to: " + callerInfo + ", Call Id: " + incomingCallId + vThreadIdInfo, datetime: getFormattedDate()};
		sendPostMessage(infoMsg);
	
		messageMsg = {type: "message", status: "PickedUp", callid: incomingCall.id, datetime: getFormattedDate()};
		sendPostMessage(messageMsg);
				
    } catch (error) {
        console.error(error);
    }

}

// Accepting an incoming call
acceptCallButton.onclick = async () => {

	goAcceptTheCall();
	
}

async function goHangUp() {
	if (btnstatus != "01") {
		return;
	}

    // end the current call
    await call1.hangUp();
	
	 document.getElementById("connectedLabel").innerHTML = "Environment is initiated!";
	 infoMsg = {type: "info", message: "Environment is initiated!", datetime: getFormattedDate()};
	 sendPostMessage(infoMsg);

	 document.getElementById("accept-call-button-image").src="hsgray.jpg";
	 document.getElementById("hangup-call-button-image").src="hsupgray.jpg";
	 btnstatus = "00";

	 actionMsg = {type: "action", command: "idle", datetime: getFormattedDate()};
	 sendPostMessage(actionMsg);

	 document.getElementById("transfer-call-button-image").src="hstgray.jpg";
	 transferCallButton.disabled = true;
	 transferTargetPhone.disabled = true;

	 actionMsg = {type: "action", command: "transferoff", datetime: getFormattedDate()};
	 sendPostMessage(actionMsg);

	 messageMsg = {type: "message", status: "HangUp", callid: incomingCall.id, datetime: getFormattedDate()};
	 sendPostMessage(messageMsg);
}

// End the current call
hangUpCallButton.addEventListener("click", async () => {

	goHangUp();

});

window.addEventListener("message", (e) => {
	if (e.data != undefined) {
		if (e.data != "") {
			if (e.data.command == "TRANSFER") {
				transferTargetPhone.value = e.data.phone;
				goMakeTransfer();
			}
			if (e.data.command == "PICKUP") {
				goAcceptTheCall();
			}
			if (e.data.command == "HANGUP") {
				goHangUp();
			}
			if (e.data.command == "MUTE") {
				goMuteCall();
			}
			if (e.data.command == "UNMUTE") {
				goUnmuteCall();
			}
		}
	}
});

async function goMuteCall() {

	await call1.mute();

	document.getElementById("connectedLabel").innerHTML = "Call is currently muted.";
	infoMsg = {type: "info", message: "Call is currently muted.", datetime: getFormattedDate()};
	sendPostMessage(infoMsg);
}

async function goUnmuteCall() {

	await call1.unmute();

	document.getElementById("connectedLabel").innerHTML = "Call is currently unmuted.";
	infoMsg = {type: "info", message: "Call is currently unmuted.", datetime: getFormattedDate()};
	sendPostMessage(infoMsg);
}

// transfer current call
transferCallButton.addEventListener("click", async () => {

	goMakeTransfer();
	
});

async function goMakeTransfer() {
	document.getElementById("connectedLabel").innerHTML = "Initiating transfer. Please wait..";
	infoMsg = {type: "info", message: "Initiating transfer. Please wait..", datetime: getFormattedDate()};
	sendPostMessage(infoMsg);

	if (transferTargetPhone.value.indexOf("+") == 0) {

		await call1.hold();
								
		// pstn
		const pstnCallee = { phoneNumber: transferTargetPhone.value }

		const transfer = callTransferApi1.transfer({targetParticipant: pstnCallee});
		transfer.on('stateChanged', async () => {
			
			document.getElementById("connectedLabel").innerHTML = "Transfer state: " + transfer.state;
			infoMsg = {type: "info", message: "Transfer state: " + transfer.state, datetime: getFormattedDate()};
			sendPostMessage(infoMsg);
		
			if (transfer.state === 'Failed') {

				document.getElementById("connectedLabel").innerHTML = "Transfer failed.";
				infoMsg = {type: "info", message: "Transfer failed.", datetime: getFormattedDate()};
				sendPostMessage(infoMsg);

				document.getElementById("transfer-target-phone").value = "";
				document.getElementById("transfer-call-button-image").src="hstgray.jpg";
				transferCallButton.disabled = true;
				transferTargetPhone.disabled = true;

				actionMsg = {type: "action", command: "transferoff", datetime: getFormattedDate()};
				sendPostMessage(actionMsg);
			}
									
			if (transfer.state === 'Transferred') {
				document.getElementById("connectedLabel").innerHTML = "Call transfered to: " + transferTargetPhone.value;
				infoMsg = {type: "info", message: "Call transfered to: " + transferTargetPhone.value, datetime: getFormattedDate()};
				sendPostMessage(infoMsg);
	
				// end the calls
				await call1.hangUp();
				
				document.getElementById("transfer-target-phone").value = "";
				document.getElementById("transfer-call-button-image").src="hstgray.jpg";
				transferCallButton.disabled = true;
				transferTargetPhone.disabled = true;

				messageMsg = {type: "message", status: "Transferred", callid: incomingCall.id, datetime: getFormattedDate()};
				sendPostMessage(messageMsg);
						   
			}
		});
	}
	else{

		var myHeaders = new Headers();
		myHeaders.append("Content-Type", "application/json");

		var requestOptions = {
			method: 'GET',
			headers: myHeaders,
			redirect: 'follow'
	  	};

		var url = "https://waframegettoken.azurewebsites.net/getAppToken";
		//var url = "https://wagettoken.azurewebsites.net/getAppToken";

		var vAppToken = "n/a";
		try {
			const rawResponse = await fetch(url, requestOptions);
			const content = await rawResponse.json();
			vAppToken = content.result;
			document.getElementById("connectedLabel").innerHTML = "Token acquired.";
			infoMsg = {type: "info", message: "Token acquired.", datetime: getFormattedDate()};
			sendPostMessage(infoMsg);			
		}
		catch (error) {
			console.error(error);
			document.getElementById("connectedLabel").innerHTML = error;
			infoMsg = {type: "info", message: error, datetime: getFormattedDate()};
			sendPostMessage(infoMsg);

			vAppToken = "n/a";
	 	}

		if (vAppToken != undefined) {
			if (vAppToken != "n/a") {

				document.getElementById("connectedLabel").innerHTML = "Creating transfering group.";
				infoMsg = {type: "info", message: "Creating transfering group.", datetime: getFormattedDate()};
				sendPostMessage(infoMsg);

				var url2 = "https://waframegettoken.azurewebsites.net/getThreadId?appToken=" + vAppToken + "&userid=" + TeamsUserId;
				//var url2 = "https://wagettoken.azurewebsites.net/getThreadId?appToken=" + vAppToken + "&userid=" + TeamsUserId;

				var vThreadId = "n/a";
				try {
					const rawResponse2 = await fetch(url2, requestOptions);
					const content2 = await rawResponse2.json();
					vThreadId = content2.result;
					document.getElementById("connectedLabel").innerHTML = "Group created.";
					infoMsg = {type: "info", message: "Group created.", datetime: getFormattedDate()};
					sendPostMessage(infoMsg);
				}
				catch (error) {
					console.error(error);
					document.getElementById("connectedLabel").innerHTML = error;
					infoMsg = {type: "info", message: error, datetime: getFormattedDate()};
					sendPostMessage(infoMsg);
					vThreadId = "n/a";
				}

				if (vThreadId != undefined) {
					if (vThreadId != "n/a") {

						await call1.hold();

						document.getElementById("connectedLabel").innerHTML = "Calling " + transferTargetPhone.value + ", ThreadId: " + vThreadId;
						infoMsg = {type: "info", message: "Calling " + transferTargetPhone.value + ", ThreadId: " + vThreadId, datetime: getFormattedDate()};
						sendPostMessage(infoMsg);
	
						// email
						document.getElementById("connectedLabel").innerHTML = "Getting user id.";
						infoMsg = {type: "info", message: "Getting user id.", datetime: getFormattedDate()};
						sendPostMessage(infoMsg);

						var url3 = "https://waframegettoken.azurewebsites.net/getUserId?appToken=" + vAppToken + "&useremail=" + transferTargetPhone.value;
						//var url3 = "https://wagettoken.azurewebsites.net/getUserId?appToken=" + vAppToken + "&useremail=" + transferTargetPhone.value;
															
						var vUserId = "n/a";
						try {
							const rawResponse3 = await fetch(url3, requestOptions);
							const content3 = await rawResponse3.json();
							vUserId = content3.result;
							document.getElementById("connectedLabel").innerHTML = "User id found.";
							infoMsg = {type: "info", message: "User id found.", datetime: getFormattedDate()};
							sendPostMessage(infoMsg);
			
							const userCallee = { microsoftTeamsUserId: vUserId };
							//call2 = callAgent.startCall([userCallee], { threadId: vThreadId });
					
							const transfer = callTransferApi1.transfer({targetParticipant: userCallee});
							transfer.on('stateChanged', async () => {
							
								document.getElementById("connectedLabel").innerHTML = "Transfer state: " + transfer.state;
								infoMsg = {type: "info", message: "Transfer state: " + transfer.state, datetime: getFormattedDate()};
								sendPostMessage(infoMsg);
	
								if (transfer.state === 'Failed') {
									document.getElementById("connectedLabel").innerHTML = "Transfer failed.";
									infoMsg = {type: "info", message: "Transfer failed.", datetime: getFormattedDate()};
									sendPostMessage(infoMsg);
									document.getElementById("transfer-target-phone").value = "";
									document.getElementById("transfer-call-button-image").src="hstgray.jpg";
									transferCallButton.disabled = true;
									transferTargetPhone.disabled = true;
								}
													
								if (transfer.state === 'Transferred') {
									document.getElementById("connectedLabel").innerHTML = "Call transfered to: " + transferTargetPhone.value;
									infoMsg = {type: "info", message: "Call transfered to: " + transferTargetPhone.value, datetime: getFormattedDate()};
									sendPostMessage(infoMsg);

									// end the calls
									await call1.hangUp();											
							
									document.getElementById("transfer-target-phone").value = "";
									document.getElementById("transfer-call-button-image").src="hstgray.jpg";
									transferCallButton.disabled = true;
									transferTargetPhone.disabled = true;
										   
								}
							});
						}
						catch (error) {
							console.error(error);
							document.getElementById("connectedLabel").innerHTML = error;
							infoMsg = {type: "info", message: error, datetime: getFormattedDate()};
							sendPostMessage(infoMsg);

							vThreadId = "n/a";
						}						
					}
					else {
						document.getElementById("connectedLabel").innerHTML = "Group failed to be created.";
						infoMsg = {type: "info", message: "Group failed to be created.", datetime: getFormattedDate()};
						sendPostMessage(infoMsg);

					}
				} 
				else {
					document.getElementById("connectedLabel").innerHTML = "Group failed to be created.";
					infoMsg = {type: "info", message: "Group failed to be created.", datetime: getFormattedDate()};
					sendPostMessage(infoMsg);

				}
			}
			else {
				document.getElementById("connectedLabel").innerHTML = "Token failed to be created.";
				infoMsg = {type: "info", message: "Token failed to be created.", datetime: getFormattedDate()};
				sendPostMessage(infoMsg);
			}
		}
		else {
			document.getElementById("connectedLabel").innerHTML = "Token failed to be created.";
			infoMsg = {type: "info", message: "Token failed to be created.", datetime: getFormattedDate()};
			sendPostMessage(infoMsg);
		}
	}

}




