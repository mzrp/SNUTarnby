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
//let transferTargetPhone = document.getElementById('transfer-target-phone');
//let transferCallButton = document.getElementById('transfer-call-button');
let acceptCallButton = document.getElementById('accept-call-button');
let hangUpCallButton = document.getElementById('hangup-call-button');

//transferCallButton.disabled = true;
//transferTargetPhone.disabled = true;	

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

		//transferCallButton.disabled = true;
		//transferTargetPhone.disabled = true;

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
						
						//transferCallButton.disabled = true;
						//transferTargetPhone.disabled = true;		
						
						document.getElementById("connectedLabel").innerHTML = "Environment is initiated!";

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
						callToShow += ", " + callPhone;
					}
	
					// Get information about caller
					callerInfo = callToShow + " (" + callKind + ")";
	
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
	///*
	document.getElementById("connectedLabel").innerHTML = "Initiating environment, please wait...";
	ACSToken = urlParams.get('token');
	ACSTokenExpires = urlParams.get('expires');

	TeamsUserEmail = urlParams.get('email');
	TeamsUserId = urlParams.get('userid');
	TeamsUserName = urlParams.get('name');
	TeamsToken = urlParams.get('teamstoken');

    CreateInstance();
	//*/
}

// Accepting an incoming call
acceptCallButton.onclick = async () => {

	if (btnstatus != "10") {
		return;
	}

    try {
        call1 = await incomingCall.accept();

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

		//transferCallButton.disabled = false;
		//transferTargetPhone.disabled = false;
		
		document.getElementById("connectedLabel").innerHTML = "Connected to: " + callerInfo + ", Call Id: " + incomingCallId;

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

	 //transferCallButton.disabled = true;
	 //transferTargetPhone.disabled = true;

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

/*
transferCallButton.addEventListener("click", async () => {

	document.getElementById("connectedLabel").innerHTML = "Initiating transfer. Please wait..";

	var myHeaders = new Headers();
	myHeaders.append("Content-Type", "application/json");

	var requestOptions = {
		method: 'GET',
		headers: myHeaders,
		redirect: 'follow'
	  };

	var url = "https://wagettoken.azurewebsites.net/getAppToken";

	const rawResponse = await fetch(url, requestOptions);
	const content = await rawResponse.json();

	var vAppToken = content.result;
	//var vAppToken = TeamsToken;

	if (vAppToken != undefined) {
		if (vAppToken != "") {

			var url2 = "https://wagettoken.azurewebsites.net/getThreadId?appToken=" + vAppToken + "&userid=" + TeamsUserId;

			const rawResponse2 = await fetch(url2, requestOptions);
			const content2 = await rawResponse2.json();

			if (content2.result != undefined) {
				if (content2.result != "") {

					await call1.hold();

					var vThreadId = content2.result;
					document.getElementById("connectedLabel").innerHTML = "Calling " + transferTargetPhone.value + ", ThreadId: " + vThreadId;
					
					const pstnCallee = { phoneNumber: transferTargetPhone.value }
					call2 = callAgent.startCall([pstnCallee], { threadId: vThreadId });

					call2.on('stateChanged', async () => {
						
						console.log(`Call state changed: ${call2.state}`);
						document.getElementById("connectedLabel").innerHTML = "Call state changed: " + call2.state;

						if(call2.state === 'Connected') {

							console.log('Call started.');
							document.getElementById("connectedLabel").innerHTML = "Call to " + transferTargetPhone.value + " connected. ThreadId: " + vThreadId;

							callTransferApi2 = call2.feature(Features.Transfer);
							callTransferApi2.on('transferRequested', args => {
								console.log(`Receive transfer request: ${args.targetParticipant}`);
								args.accept();
							});

							await call2.hold();

							const transfer = callTransferApi2.transfer({ targetCallId: call1.id });
							transfer.on('stateChanged', async () => {
								document.getElementById("connectedLabel").innerHTML = "Transfer state: " + transfer.state;
								if (transfer.state === 'Transferred') {
									document.getElementById("connectedLabel").innerHTML = "Call transfered to: " + transferTargetPhone.value;

									// end the calls
    								await call1.hangUp();
									await call2.hangUp();
								}
							 });
						} 
						
						if (call2.state === 'Disconnected') {
							console.log(`Call ended, call end reason={code=${call2.callEndReason.code}, subCode=${call2.callEndReason.subCode}}`);
						}   
					});

				}
			} 

		}
	}
});
*/
