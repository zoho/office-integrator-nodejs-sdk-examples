import * as SDK from "@zoho-corp/office-integrator-sdk";
import { readFileSync, writeFileSync } from 'fs';
const __dirname = import.meta.dirname;

class EditPDF {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in your application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var editPdfParameters = new SDK.V1.EditPdfParameters();

            editPdfParameters.setUrl("https://demo.office-integrator.com/zdocs/EventForm.pdf");
            
            var documentInfo = new SDK.V1.DocumentInfo();

            //Time value used to generate unique document every time. You can replace based on your application.
            documentInfo.setDocumentId("" + new Date().getTime());
            documentInfo.setDocumentName("EventForm.pdf");

            editPdfParameters.setDocumentInfo(documentInfo);

            var userInfo = new SDK.V1.UserInfo();

            userInfo.setUserId("1000");
            userInfo.setDisplayName("Prabakaran R");

            editPdfParameters.setUserInfo(userInfo);

            var editorSettings = new SDK.V1.PdfEditorSettings();

            editorSettings.setUnit("in");
            editorSettings.setLanguage("en");

            editPdfParameters.setEditorSettings(editorSettings);

            var editorUiOptions = new SDK.V1.PdfEditorUiOptions();

            editorUiOptions.setFileMenu("show");
            editorUiOptions.setSaveButton("show");

            editPdfParameters.setUiOptions(editorUiOptions);

            var callbackSettings = new SDK.V1.CallbackSettings();
            var saveUrlParams = new Map();

            saveUrlParams.set("auth_token", "1234");
            saveUrlParams.set("id", "123131");

            var saveUrlHeaders = new Map();

            saveUrlHeaders.set("header1", "value1");
            saveUrlHeaders.set("header2", "value2");

            callbackSettings.setSaveUrlHeaders(saveUrlHeaders);
            callbackSettings.setSaveUrlParams(saveUrlParams);
            callbackSettings.setSaveFormat("pdf");
            callbackSettings.setSaveUrl("https://officeintegrator.zoho.com/v1/api/webhook/savecallback/601e12157a25e63fc4dfd4e6e00cc3da2406df2b9a1d84a903c6cfccf92c8286");

            editPdfParameters.setCallbackSettings(callbackSettings);

            var responseObject = await sdkOperations.editPdf(editPdfParameters);

            if(responseObject != null) {
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                let pdfSessionResponseObj = responseObject.object;
    
                if(pdfSessionResponseObj != null){

                    if(pdfSessionResponseObj instanceof SDK.V1.CreateDocumentResponse){
                        console.log("\nPDF Document ID - " + pdfSessionResponseObj.getDocumentId());
                        console.log("\nPDF session ID - " + pdfSessionResponseObj.getSessionId());
                        console.log("\nPDF session URL - " + pdfSessionResponseObj.getDocumentUrl());
                        console.log("\nPDF save URL - " + pdfSessionResponseObj.getSaveUrl());
                        console.log("\nPDF delete URL - " + pdfSessionResponseObj.getDocumentDeleteUrl());
                        console.log("\nPDF session delete URL - " + pdfSessionResponseObj.getSessionDeleteUrl());
                    } else if (pdfSessionResponseObj instanceof SDK.V1.InvalidConfigurationException) {
                        console.log("\nInvalid configuration exception. Exception json - ", pdfSessionResponseObj);
                    } else {
                        console.log("\nRequest not completed successfully");
                    }
                }
            }
        } catch (error) {
            console.log("\nException while running sample code", error);
        }
    }

    //Include office-integrator-sdk package in your package json and then execute this code.

    static async initializeSdk() {

        // Refer this help page for api end point domain details -  https://www.zoho.com/officeintegrator/api/v1/getting-started.html
        let environment = await new SDK.DataCenter.Production("https://api.office-integrator.com");

        let auth = new SDK.AuthBuilder()
                        .addParam("apikey", "2ae438cf864488657cc9754a27daa480") //Update this apikey with your own apikey signed up in office integrator service
                        .authenticationSchema(await new SDK.V1.Authentication().getTokenFlow())
                        .build();

        let tokens = [ auth ];

        //Sdk application log configuration
        let logger = new SDK.LogBuilder()
            .level(SDK.Levels.INFO)
            //.filePath("<file absolute path where logs would be written>") //No I18N
            .build();

        let initialize = await new SDK.InitializeBuilder();

        await initialize.environment(environment).tokens(tokens).logger(logger).initialize();

        console.log("SDK initialized successfully.");
    }
}

EditPDF.execute();