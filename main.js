const { app } = require('electron')

// define value for external program calls
const child = require('child_process').execFileSync;
const executablePath = 'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE';
// const executablePath ="C:\\GitHub\\electron-deep-linking-mac-win\\startme.bat"
const parameters = ['C:\\Visbo\\VISBO SPE\\Visbo Project Edit.xlsx'];

// logger initialize
var log4js = require("log4js");
log4js.configure({
	appenders: { everything: { type: "file", filename: "C:\\GitHub\\visbo-connect\\mydebug.log" } },
	categories: { default: { appenders: ["everything"], level: "debug" } }
  });
const logger = log4js.getLogger();

// Log both at dev console and at running node console instance
function logEverywhere(s) {
	console.log(s)
	logger.debug(s)
}

// ---------------- FUNCTIONS ----------------

// make parameter usage for simple edit
function getParameter (myParameter) {
	if (myParameter === undefined) {return undefined}

	// only the parameters after the ? are important for SPE
	const afterQuestionMarkString = myParameter.split("?")[1];
	logEverywhere ("Received from WebSite after the ? character =  " + afterQuestionMarkString)

	// clean strings from ? and substitute with / (special for spe)
	// SPE needs leading /e and final / at the end
	cleanString = afterQuestionMarkString.replace (/\&/g, "/")
	parameterForExcelAddin = '/e/' + cleanString + '/'
	return (parameterForExcelAddin)
}

// Start Program with parameters received from the web
function startExternalProgramWithParameters (parameterValue) {
		let speParameter = getParameter(parameterValue)
		logEverywhere ('Parameter for SPE = ' + speParameter)
		parameters.push(speParameter)

		logEverywhere ('START PROGRAM ' + executablePath)
		logEverywhere ('Parameters = ' + parameters)
		child(executablePath, parameters);
		logEverywhere('Program Finished !!!')
}

// create a main window for the application
function mainProgram() {
	argv = process.argv;
	logEverywhere ('argv = ' + argv)
	if (argv[1] === undefined) {
		// This an installation call
		logEverywhere('called for installation')
		if (!app.isDefaultProtocolClient('visbo-connect')) {
			// Define custom protocol handler. Deep linking works on packaged versions of the application!
			logEverywhere(' ----------------- define protocol handler --------------- ')
			app.setAsDefaultProtocolClient('visbo-connect')
		  }
	}
	else { //This is a call for starting an external program
		if (argv[2] === undefined) {
			// this the call from a web site
			logEverywhere('called from web site')
			parameterValue = argv[1]
		}
		else {
			// This is a test call npm start (First two parameters are electron and .)
			logEverywhere('called via npm for test')
			parameterValue = argv[2]
		
		}
	// start the program with parameters given by the web site or the test run
	startExternalProgramWithParameters (parameterValue)
	}
	// flush lof files and quit application
	log4js.shutdown(function() { app.quit() })
}


// This method will be called when Electron has finished
// initialization
// Some APIs can only be used after this event occurs.
// REWORK use => function
logEverywhere(' ----------------- Wait for app.on ready--------------- ')
app.on('ready', mainProgram)
