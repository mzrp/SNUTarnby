import soundfile from './TeamsRingTone.mp3'; 

// Make sure to install the necessary dependencies
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

// UI objects
let transferTargetPhone = document.getElementById('transfer-target-phone');
let transferCallButton = document.getElementById('transfer-call-button');
let acceptCallButton = document.getElementById('accept-call-button');
let hangUpCallButton = document.getElementById('hangup-call-button');

document.getElementById("transfer-call-button-image").src="hstgray.jpg";
transferCallButton.disabled = true;
transferTargetPhone.disabled = true;	

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

// Create an instance of CallClient. Initialize a CallAgent instance with a CommunicationUserCredential via created CallClient. 
async function CreateInstance() {
    try {
        const callClient = new CallClient(); 
        tokenCredential = new AzureCommunicationTokenCredential(ACSToken);
        callAgent = await callClient.createCallAgent(tokenCredential)
		
		document.getElementById("accept-call-button-image").src="hsgray.jpg";
		document.getElementById("hangup-call-button-image").src="hshugray.jpg";
		btnstatus = "00";

		document.getElementById("transfer-call-button-image").src="hstgray.jpg";
		transferCallButton.disabled = true;
		transferTargetPhone.disabled = true;

		document.getElementById("connectedLabel").innerHTML = "Environment is initiated!";
        
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
						
						document.getElementById("transfer-call-button-image").src="hstgray.jpg";
						transferCallButton.disabled = true;
						transferTargetPhone.disabled = true;	
						
						document.getElementById("connectedLabel").innerHTML = "Current call ended. Environment is initiated.";

						try {
							var infoMsg = `HangUp ${incomingCall.id} ${getFormattedDate()}`;
							console.log(infoMsg);
							var childWindow = document.getElementById("selvbetjeningfr");
							childWindow = childWindow ? childWindow.contentWindow : null;
							childWindow.postMessage(infoMsg, "*");
						}
						catch (error) {
						   console.error(error);
						}

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
					
					document.getElementById("connectedLabel").innerHTML = "Ring... Ring... Incoming call from: " + callerInfo + ", Call Id: " + incomingCallId;
								
					// play ring tune
					ringTone.load();
					ringTone.play(); 

					var infoMsg = `Presented ${incomingCall.id} ${getFormattedDate()}`;
					console.log(infoMsg);
					var childWindow = document.getElementById("selvbetjeningfr");
					childWindow = childWindow ? childWindow.contentWindow : null;
					childWindow.postMessage(infoMsg, "*");
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
	window.location.href = "https://wagettoken.azurewebsites.net/";
}

const urlParams = new URLSearchParams(queryString);
if ((urlParams == null) || (urlParams == "")) 
{
	document.getElementById("connectedLabel").innerHTML = "Acquiring token, please wait...";
	window.location.href = "https://wagettoken.azurewebsites.net/";
}
else 
{
	document.getElementById("connectedLabel").innerHTML = "Initiating environment, please wait...";
	ACSToken = urlParams.get('token');
	ACSTokenExpires = urlParams.get('expires');

	TeamsUserEmail = urlParams.get('email');
	TeamsUserId = urlParams.get('userid');
	TeamsUserName = urlParams.get('name');
	TeamsToken = urlParams.get('teamstoken');

    CreateInstance();
}

// Accepting an incoming call
acceptCallButton.onclick = async () => {

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

		document.getElementById("transfer-call-button-image").src="hstgr.jpg";
		transferCallButton.disabled = false;
		transferTargetPhone.disabled = false;

		var vThreadIdInfo = "";
		if (call1.info.threadId != undefined) {
			vThreadIdInfo =  ", ThreadId found: " + call1.info.threadId;
		}
		
		document.getElementById("connectedLabel").innerHTML = "Connected to: " + callerInfo + ", Call Id: " + incomingCallId + vThreadIdInfo;

		var infoMsg = `PickedUp ${incomingCall.id} ${getFormattedDate()}`;
		console.log(infoMsg);
		var childWindow = document.getElementById("selvbetjeningfr");
		childWindow = childWindow ? childWindow.contentWindow : null;
		childWindow.postMessage(infoMsg, "*");
				
    } catch (error) {
        console.error(error);
    }
}

// End the current call
hangUpCallButton.addEventListener("click", async () => {

	if (btnstatus != "01") {
		return;
	}

    // end the current call
    await call1.hangUp();
	
	 document.getElementById("connectedLabel").innerHTML = "Environment is initiated!";

	 document.getElementById("accept-call-button-image").src="hsgray.jpg";
	 document.getElementById("hangup-call-button-image").src="hsupgray.jpg";
	 btnstatus = "00";

	 document.getElementById("transfer-call-button-image").src="hstgray.jpg";
	 transferCallButton.disabled = true;
	 transferTargetPhone.disabled = true;

	 try {
	 	var infoMsg = `HangUp ${incomingCall.id} ${getFormattedDate()}`;
	 	console.log(infoMsg);
	 	var childWindow = document.getElementById("selvbetjeningfr");
	 	childWindow = childWindow ? childWindow.contentWindow : null;
	 	childWindow.postMessage(infoMsg, "*");
	 }
	 catch (error) {
		console.error(error);
	 }
});

// transfer current call
transferCallButton.addEventListener("click", async () => {

	document.getElementById("connectedLabel").innerHTML = "Initiating transfer. Please wait..";

	if (transferTargetPhone.value.indexOf("+") == 0) {

		await call1.hold();
								
		// pstn
		const pstnCallee = { phoneNumber: transferTargetPhone.value }

		const transfer = callTransferApi1.transfer({targetParticipant: pstnCallee});
		transfer.on('stateChanged', async () => {
			
			document.getElementById("connectedLabel").innerHTML = "Transfer state: " + transfer.state;

			if (transfer.state === 'Failed') {

				document.getElementById("transfer-target-phone").value = "";
				document.getElementById("transfer-call-button-image").src="hstgray.jpg";
				transferCallButton.disabled = true;
				transferTargetPhone.disabled = true;
			}
									
			if (transfer.state === 'Transferred') {
				document.getElementById("connectedLabel").innerHTML = "Call transfered to: " + transferTargetPhone.value;
		
				// end the calls
				await call1.hangUp();
				
				document.getElementById("transfer-target-phone").value = "";
				document.getElementById("transfer-call-button-image").src="hstgray.jpg";
				transferCallButton.disabled = true;
				transferTargetPhone.disabled = true;
						   
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

		var url = "https://wagettoken.azurewebsites.net/getAppToken";

		var vAppToken = "n/a";
		try {
			const rawResponse = await fetch(url, requestOptions);
			const content = await rawResponse.json();
			vAppToken = content.result;
			document.getElementById("connectedLabel").innerHTML = "Token acquired.";
		}
		catch (error) {
			console.error(error);
			document.getElementById("connectedLabel").innerHTML = error;
			vAppToken = "n/a";
	 	}

		if (vAppToken != undefined) {
			if (vAppToken != "n/a") {

				document.getElementById("connectedLabel").innerHTML = "Creating transfering group.";
				var url2 = "https://wagettoken.azurewebsites.net/getThreadId?appToken=" + vAppToken + "&userid=" + TeamsUserId;

				var vThreadId = "n/a";
				try {
					const rawResponse2 = await fetch(url2, requestOptions);
					const content2 = await rawResponse2.json();
					vThreadId = content2.result;
					document.getElementById("connectedLabel").innerHTML = "Group created.";
				}
				catch (error) {
					console.error(error);
					document.getElementById("connectedLabel").innerHTML = error;
					vThreadId = "n/a";
				}

				if (vThreadId != undefined) {
					if (vThreadId != "n/a") {

						await call1.hold();

						document.getElementById("connectedLabel").innerHTML = "Calling " + transferTargetPhone.value + ", ThreadId: " + vThreadId;
					
						// email
						document.getElementById("connectedLabel").innerHTML = "Getting user id.";
						var url3 = "https://wagettoken.azurewebsites.net/getUserId?appToken=" + vAppToken + "&useremail=" + transferTargetPhone.value;
															
						var vUserId = "n/a";
						try {
							const rawResponse3 = await fetch(url3, requestOptions);
							const content3 = await rawResponse3.json();
							vUserId = content3.result;
							document.getElementById("connectedLabel").innerHTML = "User id found.";
											
							const userCallee = { microsoftTeamsUserId: vUserId };
							//call2 = callAgent.startCall([userCallee], { threadId: vThreadId });
					
							const transfer = callTransferApi1.transfer({targetParticipant: userCallee});
							transfer.on('stateChanged', async () => {
							
								document.getElementById("connectedLabel").innerHTML = "Transfer state: " + transfer.state;

								if (transfer.state === 'Failed') {

									document.getElementById("transfer-target-phone").value = "";
									document.getElementById("transfer-call-button-image").src="hstgray.jpg";
									transferCallButton.disabled = true;
									transferTargetPhone.disabled = true;
								}
													
								if (transfer.state === 'Transferred') {
									document.getElementById("connectedLabel").innerHTML = "Call transfered to: " + transferTargetPhone.value;
						
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
							vThreadId = "n/a";
						}						
					}
					else {
						document.getElementById("connectedLabel").innerHTML = "Group failed to be created.";
					}
				} 
				else {
					document.getElementById("connectedLabel").innerHTML = "Group failed to be created.";
				}
			}
			else {
				document.getElementById("connectedLabel").innerHTML = "Token failed to be created.";
			}
		}
		else {
			document.getElementById("connectedLabel").innerHTML = "Token failed to be created.";
		}
	}
});


