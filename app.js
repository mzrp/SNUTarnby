import soundfile from './TeamsRingTone.mp3'; 
  
const { CallClient, VideoStreamRenderer, LocalVideoStream } = require('@azure/communication-calling');
const { AzureCommunicationTokenCredential } = require('@azure/communication-common');
 
const { AzureLogger, setLogLevel, createClientLogger } = require("@azure/logger");

import { Features} from "@azure/communication-calling";
import { apply } from 'file-loader';

var JSZip = require("jszip");

// Set the log level and output
const _logger = createClientLogger('ACS');
const _log = [];
const _maxLogSize = 100000;
setLogLevel('verbose');
AzureLogger.log = (...args) => {
    console.log(...args);
    try {
        if (_log.length > _maxLogSize) {
            _log.splice(0, _maxLogSize / 10);
        }
        if (Array.isArray(args)) {
            for (let i = 0; i < args.length; i++) {
                const s = args[i].toString();
                _log.push(s);
            }
        }
    } catch (error) {
        console.log('AzureLogger.log error', error);
    }	
};

function SaveLogEntry(logentry) {
	console.log(logentry);
    try {
        if (_log.length > _maxLogSize) {
            _log.splice(0, _maxLogSize / 10);
        }        
		if (Array.isArray(logentry)) {
            for (let i = 0; i < logentry.length; i++) {
                const s = logentry[i].toString();
                _log.push("#RPCODELOG#: " + s);
            }
        }
		else {
			if (logentry != null && typeof logentry == 'object') {
				_log.push("#RPCODELOG#: " + JSON.stringify(logentry));			
			}
			else {
				_log.push("#RPCODELOG#: " + logentry);
			}
		}
    } catch (error) {
        console.log('AzureLogger.log error', error);
    }	
}

function ACSDownloadLog() {
    const a = document.createElement("a");
    const data = _log.join('\r\n');
    const blob = new Blob([data], { type: "octet/stream" });

	var zip = new JSZip();
	zip.file("log.txt", blob);
	zip.generateAsync({ type:"blob", compression: "DEFLATE", compressionOptions: {level: 9} })
		.then(function(content) {
		    const url = window.URL.createObjectURL(content);
    		a.href = url;
    		a.download = 'log.zip';
    		a.click();
    		window.URL.revokeObjectURL(url);	
	});

}

// UI objects
let transferTargetPhone = document.getElementById('transfer-target-phone');
let transferCallButton = document.getElementById('transfer-call-button');
let acceptCallButton = document.getElementById('accept-call-button');
let hangUpCallButton = document.getElementById('hangup-call-button');
let muteunmuteCallButton = document.getElementById('muteunmute-call-button');

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

let callClient;
let callAgent;
let deviceManager;

let ACSToken;
let ACSTokenExpires;
let TeamsUserName;
let TeamsUserId;
let TeamsToken;
let TeamsUserEmail;
let TeamsUserAccount;
let tokenCredential;

let call1;
let call2;
let callTransferApi1;
let callTransferApi2;

let incomingCall;
let incomingCallId;
let callerInfo;
let ringTone;
var btnstatus = "00";

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

		// send message to the iframe window
		//var childWindow = document.getElementById("selvbetjeningfr");
		//childWindow = childWindow ? childWindow.contentWindow : null;
		//childWindow.postMessage(infoMsg, "*");

		// send message to the parent window
		var parentWindow = window.parent;
		parentWindow.postMessage(infoMsg, "*");

		SaveLogEntry(JSON.stringify(infoMsg));
	}
	catch (error) {
		SaveLogEntry(error.message);
	}

}

//const fetchTokenFromMyServerForUser = async function (username) {
async function fetchTokenFromMyServerForUser(username) {

	SaveLogEntry("[REFRESHTOKEN] fetchTokenFromMyServerForUser arrived.");

	var urlGetRefreshToken = "https://taarnby-henven-gettoken.azurewebsites.net/getTokenForTeamsUser?useraccount=" + TeamsUserAccount;

	SaveLogEntry("[REFRESHTOKEN] url: " + urlGetRefreshToken);

	// Exchange the Azure AD access token of the Teams User for a Communication Identity access token
	try {
		var myTokenHeaders = new Headers();
		myTokenHeaders.append("Content-Type", "application/json");
		var requestTokenOptions = {
			method: 'GET',
			headers: myTokenHeaders,
			redirect: 'follow'
	  	};
		const newTokenResponse = await fetch(urlGetRefreshToken, requestTokenOptions);
		const newTokenResponseJson = await newTokenResponse.json();

		SaveLogEntry("[REFRESHTOKEN] response ok, New token: " + newTokenResponseJson.result);	

		if (newTokenResponseJson.result != null) {
			if (newTokenResponseJson.result.indexOf("RefreshTokenError") == -1) {
				// refresh token if no call is in progress
				setTimeout(goRefreshTokenNow.bind(null, newTokenResponseJson.result), 3000);
			}
			else {
				messageMsg = {type: "message", status: "RefreshToken", text: "Error Refreshing Token. Please Log In.", datetime: getFormattedDate()};
				sendPostMessage(messageMsg);
			}
		}
		else {
			messageMsg = {type: "message", status: "RefreshToken", text: "Error Refreshing Token. Please Log In.", datetime: getFormattedDate()};
			sendPostMessage(messageMsg);
		}

	}
	catch (error) {
		SaveLogEntry("[REFRESHTOKEN] fetch failed: " + error.message);
	}

	SaveLogEntry("[REFRESHTOKEN] fetch finished.");
}

function goRefreshTokenNow(newRefreshedToken) {
	
	if (btnstatus == "00") {

		try {
			ACSToken = newRefreshedToken;
			SaveLogEntry("[REFRESHTOKEN] refreshed token assigned.");	
		}
		catch (errordisp) {
			SaveLogEntry("[REFRESHTOKEN] ISS1.");
		}		

		try {
			tokenCredential.dispose();
			callAgent.dispose();
			SaveLogEntry("[REFRESHTOKEN] disposing done.");	
		}
		catch (errordisp) {
			SaveLogEntry("[REFRESHTOKEN] ISS2.");
		}

		try {
			callClient = null;
			deviceManager = null;
			callAgent = null;
			SaveLogEntry("[REFRESHTOKEN] vars nulled.");		
		}
		catch (errordisp) {
			SaveLogEntry("[REFRESHTOKEN] ISS3.");
		}
		
		CreateInstance(false);

		SaveLogEntry("[REFRESHTOKEN] New token activated.");

		messageMsg = {type: "message", status: "RefreshToken", text: "Token Refreshed.", datetime: getFormattedDate()};
		sendPostMessage(messageMsg);
	}
	else {
		// refresh token if no call is in progress
		setTimeout(goRefreshTokenNow.bind(null, newRefreshedToken), 3000);
	}

}

// Create an instance of CallClient. Initialize a CallAgent instance with a CommunicationUserCredential via created CallClient. 
async function CreateInstance(firsttime) {
    try {       
		// short-lived
		tokenCredential = new AzureCommunicationTokenCredential(ACSToken);

		/*
		tokenCredential = new AzureCommunicationTokenCredential({
            tokenRefresher: async () => fetchTokenFromMyServerForUser(TeamsUserEmail),
			refreshProactively: true,
            token: ACSToken
        });
		*/

		// Initiate call agent
		callClient = new CallClient({ _logger }); 
		callAgent = await callClient.createCallAgent(tokenCredential);

		// refresh token in 30 min
		setTimeout(fetchTokenFromMyServerForUser.bind(null, TeamsUserEmail), 1800000);

		// Set up a audio device to use.
		deviceManager = await callClient.getDeviceManager();
		await deviceManager.askDevicePermission({ audio: true });					
				
		if (btnstatus == "00") {

			document.getElementById("accept-call-button-image").src="hsgray.jpg";
			document.getElementById("hangup-call-button-image").src="hshugray.jpg";
			document.getElementById("muteunmute-call-button-image").src="mutegray.jpg";
			btnstatus = "00";		

			actionMsg = {type: "action", command: "idle", datetime: getFormattedDate()};
			sendPostMessage(actionMsg);

			document.getElementById("transfer-call-button-image").src="hstgray.jpg";
			transferCallButton.disabled = true;
			transferTargetPhone.disabled = true;

			actionMsg = {type: "action", command: "transferoff", datetime: getFormattedDate()};
			sendPostMessage(actionMsg);
		}

		if (firsttime == true) {
			document.getElementById("connectedLabel").innerHTML = "Environment is initiated!";
			infoMsg = {type: "info", message: "Environment is initiated!", datetime: getFormattedDate()};
			sendPostMessage(infoMsg);
		}
		else {
			document.getElementById("connectedLabel").innerHTML = "Token refreshed. Environment is reinitiated!";
			infoMsg = {type: "info", message: "Environment is initiated!", datetime: getFormattedDate()};
			sendPostMessage(infoMsg);
		}
		        		
        // Listen for an incoming call to accept.
        callAgent.on('incomingCall', async (args) => {
            try {
					incomingCall = args.incomingCall;	

					incomingCall.on('callEnded', args => {
						SaveLogEntry(args.callEndReason);

						document.getElementById("accept-call-button-image").src="hsgray.jpg";
						document.getElementById("hangup-call-button-image").src="hshugray.jpg";
						document.getElementById("muteunmute-call-button-image").src="mutegray.jpg";
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

						// go set presence to busy/inacall
						SetPresence("Available", "Available");	 

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
							if (callPhone != null) {
								callToShow += ", " + callPhone;
							}
						}
					}
					if (callToShow != null) {
						if (callToShow.indexOf(", ") == 0) {
							callToShow = callToShow.replace(", ", "");
						}						
					}
					else {
						callToShow = callPhone;
					}					
					
					// Get information about caller
					callerInfo = callToShow;// + " (" + callKind + ")";

					if (callerInfo == null) {
						callerInfo = "Queue call";
					}
			
					if (callerInfo == "") {
						callerInfo = "Queue call";
					}					
	
					document.getElementById("accept-call-button-image").src="hsgr.jpg";
					document.getElementById("hangup-call-button-image").src="hshugray.jpg";
					document.getElementById("muteunmute-call-button-image").src="mutegray.jpg";
					btnstatus = "10";
					
					actionMsg = {type: "action", command: "callringing", caller: callerInfo, datetime: getFormattedDate()};
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
						SaveLogEntry(error.message);
					}

					messageMsg = {type: "message", status: "Presented", callid: incomingCall.id, datetime: getFormattedDate()};
					sendPostMessage(messageMsg);

            } catch (error) {
                SaveLogEntry(error.message);
            }
        });
        //initializeCallAgentButton.disabled = true;
    } catch(error) {
        SaveLogEntry(error.message);
    }
}

const queryString = window.location.search;
if ((queryString == null) || (queryString == "")) 
{
	document.getElementById("connectedLabel").innerHTML = "Acquiring token, please wait...";
	infoMsg = {type: "info", message: "Acquiring token, please wait...", datetime: getFormattedDate()};
	sendPostMessage(infoMsg);

	window.location.href = "https://taarnby-henven-gettoken.azurewebsites.net";
}

const urlParams = new URLSearchParams(queryString);
if ((urlParams == null) || (urlParams == "")) 
{
	document.getElementById("connectedLabel").innerHTML = "Acquiring token, please wait...";
	infoMsg = {type: "info", message: "Acquiring token, please wait...", datetime: getFormattedDate()};
	sendPostMessage(infoMsg);

	window.location.href = "https://taarnby-henven-gettoken.azurewebsites.net/";
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
	TeamsUserAccount = unescape(urlParams.get('teamsaccount'));
	//TeamsUserAccount = JSON.parse(unescape(urlParams.get('teamsaccount')));

	document.getElementById("internalLabel").innerHTML = "<a href='https://taarnbyrcswebanswer.z6.web.core.windows.net/Api/Phonebook.html?teamsuserid=" + TeamsUserId + "&teamstoken=" + TeamsToken + "' target='_blank'>Phonebook</a>";

    CreateInstance(true);
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
				if (callPhone != null) {
					callToShow += ", " + callPhone;
				}
			}
		}
		if (callToShow != null) {
			if (callToShow.indexOf(", ") == 0) {
				callToShow = callToShow.replace(", ", "");
			}						
		}
		else {
			callToShow = callPhone;
		}		

		callerInfo = callToShow;// + " (" + callKind + ")";

		if (callerInfo == null) {
			callerInfo = "Queue call";
		}

		if (callerInfo == "") {
			callerInfo = "Queue call";
		}		

		callTransferApi1 = call1.feature(Features.Transfer);
		callTransferApi1.on('transferRequested', args => {
			SaveLogEntry(`Receive transfer request: ${args.targetParticipant}`);
			args.accept();
		});

		// stop ring tune
		ringTone.pause();
		
		document.getElementById("accept-call-button-image").src="hsgray.jpg";
		document.getElementById("hangup-call-button-image").src="hshured.jpg";
		document.getElementById("muteunmute-call-button-image").src="muteblue.jpg";
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
	
		messageMsg = {type: "message", status: "PickedUp", caller: callerInfo, callid: incomingCall.id, datetime: getFormattedDate()};
		sendPostMessage(messageMsg);

		// go set presence to busy/inacall
		SetPresence("Busy", "InACall");
				
    } catch (error) {
        SaveLogEntry(error.message);
    }

}

async function SetPresence(vPresenceAvailability, vPresenceActivity) {

	var myHeaders = new Headers();
	myHeaders.append("Content-Type", "application/json");

	var requestOptions = {
		method: 'GET',
		headers: myHeaders,
		redirect: 'follow'
	  };

	var url = "https://taarnby-henven-gettoken.azurewebsites.net/getAppToken";	

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
		SaveLogEntry(error.message);
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

			var url2 = "https://taarnby-henven-gettoken.azurewebsites.net/setPresence?appToken=" + vAppToken + "&userid=" + TeamsUserId + "&presAvailability=" + vPresenceAvailability + "&presActivity=" + vPresenceActivity;

			var vThreadId = "n/a";
			try {
				const rawResponse2 = await fetch(url2, requestOptions);
				const content2 = await rawResponse2.json();
				vThreadId = content2.result;
				document.getElementById("connectedLabel").innerHTML = "Presence updated to " + vPresenceAvailability + " (" + vPresenceActivity + ")";
				infoMsg = {type: "info", message: "Presence updated to " + vPresenceAvailability + " (" + vPresenceActivity + ")", datetime: getFormattedDate()};
				sendPostMessage(infoMsg);
			}
			catch (error) {
				SaveLogEntry(error.message);
				document.getElementById("connectedLabel").innerHTML = error;
				infoMsg = {type: "info", message: error, datetime: getFormattedDate()};
				sendPostMessage(infoMsg);
				vThreadId = "n/a";
			}
		}
	}

}

// Muting an incoming call
muteunmuteCallButton.onclick = async () => {

	//alert(document.getElementById("muteunmute-call-button-image").src);

	if (document.getElementById("muteunmute-call-button-image").src == "https://taarnbyrcswebanswer.z6.web.core.windows.net/muteblue.jpg") {
		goMuteCall();
	}

	if (document.getElementById("muteunmute-call-button-image").src == "https://taarnbyrcswebanswer.z6.web.core.windows.net/mutered.jpg") {
		goUnmuteCall();
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
	 document.getElementById("hangup-call-button-image").src="hshugray.jpg";
	 document.getElementById("muteunmute-call-button-image").src="mutegray.jpg";
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

	 // go set presence to busy/inacall
	 SetPresence("Available", "Available");	 
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
				SaveLogEntry("[SELVBETMSG] PICKUP arrived.");
				goAcceptTheCall();
			}
			if (e.data.command == "HANGUP") {
				SaveLogEntry("[SELVBETMSG] HANGUP arrived.");
				goHangUp();
			}
			if (e.data.command == "MUTE") {
				SaveLogEntry("[SELVBETMSG] MUTE arrived.");
				goMuteCall();
			}
			if (e.data.command == "UNMUTE") {
				SaveLogEntry("[SELVBETMSG] UNMUTE arrived.");
				goUnmuteCall();
			}
			if (e.data.command == "ACSDOWNLOADLOG") {
				SaveLogEntry("[SELVBETMSG] ACSDOWNLOADLOG arrived.");
				ACSDownloadLog();
			}
		}
	}
});

async function goMuteCall() {

	await call1.mute();

	document.getElementById("muteunmute-call-button-image").src = "mutered.jpg"
	document.getElementById("connectedLabel").innerHTML = "Call is currently muted.";
	infoMsg = {type: "info", message: "Call is currently muted.", datetime: getFormattedDate()};
	sendPostMessage(infoMsg);
}

async function goUnmuteCall() {

	await call1.unmute();

	document.getElementById("muteunmute-call-button-image").src = "muteblue.jpg"
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

				// go set presence to busy/inacall
				SetPresence("Available", "Available");				
						   
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

		var url = "https://taarnby-henven-gettoken.azurewebsites.net/getAppToken";		

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
			SaveLogEntry(error.message);
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

				var url2 = "https://taarnby-henven-gettoken.azurewebsites.net/getThreadId?appToken=" + vAppToken + "&userid=" + TeamsUserId;

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
					SaveLogEntry(error.message);
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

						var url3 = "https://taarnby-henven-gettoken.azurewebsites.net/getUserId?appToken=" + vAppToken + "&useremail=" + transferTargetPhone.value;															

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

									// go set presence to busy/inacall
									SetPresence("Available", "Available");									
										   
								}
							});
						}
						catch (error) {
							SaveLogEntry(error.message);
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




