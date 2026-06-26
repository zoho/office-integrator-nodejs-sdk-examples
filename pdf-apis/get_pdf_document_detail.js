import * as SDK from "@zoho-corp/office-integrator-sdk";

class GetPdfDocumentDetail {

    static async execute() {
        
        //Initializing SDK once is enough. Calling here since code sample will be tested standalone. 
        //You can place SDK initializer code in your application and call once while your application start-up. 
        await this.initializeSdk();

        try {
            var sdkOperations = new SDK.V1.V1Operations();
            var editPdfParameters = new SDK.V1.EditPdfParameters();
            
            editPdfParameters.setUrl("https://demo.office-integrator.com/zdocs/EventForm.pdf");

            var responseObject = await sdkOperations.editPdf(editPdfParameters);

            var sessionObject = responseObject.object;

            if( sessionObject instanceof SDK.V1.CreateDocumentResponse ) {
                var documentId = sessionObject.getDocumentId();

                console.log("\nPDF session created to demonstrate get document details api. Created session ID - ", documentId);

                var responseObject = await sdkOperations.getPdfDocumentInfo(documentId);

                if(responseObject != null) {
                    //Get the status code from response
                    console.log("\nStatus Code: " + responseObject.statusCode);
        
                    //Get the api response object from responseObject
                    let documentMetaObj = responseObject.object;
        
                    if(documentMetaObj != null){
                        //TODO: Need to fix object type issue
                        if(documentMetaObj instanceof SDK.V1.DocumentMeta ){
                            console.log("\nDocument ID - " + documentMetaObj.getDocumentId());
                            console.log("\nSession Expires On - " + documentMetaObj.getExpiresOn());
                            console.log("\nSession Expires On MS - " + documentMetaObj.getExpiresOnMs());
                            console.log("\nDocument Name - " + documentMetaObj.getDocumentName());
                            console.log("\nDocument Type - " + documentMetaObj.getDocumentType());
                            console.log("\nDocument Created Time - " + documentMetaObj.getCreatedTime());
                            console.log("\nDocument Created Time MS - " + documentMetaObj.getCreatedTimeMs());
                        } else if (documentMetaObj instanceof SDK.V1.InvalidConfigurationException) {
                            console.log("\nInvalid configuration exception. Exception json - ", documentMetaObj);
                        } else {
                            console.log("\nRequest not completed successfully");
                        }
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

GetPdfDocumentDetail.execute();