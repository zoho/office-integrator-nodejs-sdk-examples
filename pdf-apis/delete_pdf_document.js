import * as SDK from "@zoho-corp/office-integrator-sdk";

class DeletePdfDocument {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in your application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var editPdfParameters = new SDK.V1.EditPdfParameters();
                        
            editPdfParameters.setUrl("https://demo.office-integrator.com/zdocs/EventForm.pdf");

            var responseObject = await sdkOperations.editPdf(editPdfParameters);

            var documentId = responseObject.object.getDocumentId();

            console.log("\nPDF id to be deleted - ", documentId);

            responseObject = await sdkOperations.deletePdfDocument(documentId);

            if(responseObject != null) {
                //Get the status code from response
                console.log("\nStatus Code: " + responseObject.statusCode);
    
                //Get the api response object from responseObject
                let pdfDeleteResponseObject = responseObject.object;
    
                if(pdfDeleteResponseObject != null){
                    if(pdfDeleteResponseObject instanceof SDK.V1.DocumentDeleteSuccessResponse){
                        console.log("\nDocument delete status - " + pdfDeleteResponseObject.getDocumentDeleted());
                    } else if (pdfDeleteResponseObject instanceof SDK.V1.InvalidConfigurationException) {
                        console.log("\nInvalid configuration exception. Exception json - ", pdfDeleteResponseObject);
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

DeletePdfDocument.execute();