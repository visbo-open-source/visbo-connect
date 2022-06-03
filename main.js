const { app } = require('electron')

// define value for external program calls
const child = require('child_process').execFileSync;
const executablePath = 'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE';
// const executablePath ="C:\\GitHub\\electron-deep-linking-mac-win\\startme.bat"
const parameters = ['C:\\Visbo\\Software\\Projectboard.xlsx'];

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
// start an external program with parameters
function startExternalProgram (executablePath, parameters) {
	logEverywhere ('--- START PROGRAM ---')
	logEverywhere ('parameters = ' + parameters)
    child(executablePath, parameters);
    logEverywhere('program finished = ')
}

// make paramter usage for simple edit
function getParameter (myParameter) {
	if (myParameter === undefined) {return undefined}
  	logEverywhere ('myParameter = ' + myParameter)
	const afterQuestionMarkString = myParameter.split("?")[1];
	logEverywhere ("afterQuestionMarkStringRight =  " + afterQuestionMarkString)

	// rework code (special for spe)
	parameterForExcelAddin = '/"' + afterQuestionMarkString + '"'
  	// cleanString = parameterForExcelAddin.replace (/\&/g, "ZZZ")
	cleanString = parameterForExcelAddin
	logEverywhere ( 'return for excel = ' + cleanString)
	return (cleanString)
}

// Start Program with parameter received from the web
function startExternalProgramWithParameters (parameterValue) {
		logEverywhere ('parameterValue = ' + parameterValue)
		let speParameter = getParameter(parameterValue)
		logEverywhere ('speParameter = ' + speParameter)

		parameters.push(speParameter)
		startExternalProgram (executablePath, parameters)
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
	else {
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
	startExternalProgramWithParameters (parameterValue)
	}
	app.quit()
}


// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
// REWORK use => function
logEverywhere(' ----------------- wait for app.on ready NEW --------------- ')
app.on('ready', mainProgram)

