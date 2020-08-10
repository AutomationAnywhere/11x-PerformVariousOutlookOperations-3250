package com.automationanywhere.botcommand.twiliomessenger.commands;

import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.DataType;
import com.twilio.Twilio;
import com.twilio.rest.api.v2010.account.Call;
import com.twilio.type.PhoneNumber;
import static com.automationanywhere.commandsdk.model.AttributeType.CREDENTIAL;
import static com.automationanywhere.commandsdk.model.AttributeType.TEXT;
import static com.automationanywhere.commandsdk.model.DataType.STRING;
import com.automationanywhere.core.security.SecureString;
import com.automationanywhere.botcommand.data.impl.StringValue;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPkg adds required information to be displayed on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        label = "[[MakeCall.label]]",
        name = "[[MakeCall.name]]",
        description = "[[MakeCall.description]]",
        icon ="twilio.svg",
        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_type = DataType.STRING,
        return_label = "[[MakeCall.return_label]]",
        node_label = "[[MakeCall.node_label]]",
        return_required = true)


public class MakeCall {
    //Identify the entry point for the action. Returns a StringValue because the return type is String.
    @Execute
    public StringValue PlaceCall(

            //Idx 1 would be displayed first, with a credential box for entering the value.
            @Idx(index = "1", type = CREDENTIAL)
            //UI labels.
            @Pkg(label = "[[MakeCall.SID.label]]",   default_value_type = STRING, description = "[[MakeCall.SID.description]]")
            //Ensure that a validation error is thrown when the value is null.
            @NotEmpty SecureString authSID,

            @Idx(index = "2", type = CREDENTIAL)
            @Pkg(label = "[[MakeCall.AuthToken.label]]",   default_value_type = STRING, description = "[[MakeCall.AuthToken.description]]")
            @NotEmpty SecureString authToken,

            @Idx(index = "3", type = TEXT)
            @Pkg(label = "[[MakeCall.SenderNumber.label]]", description = "[[MakeCall.SenderNumber.description]]")
            @NotEmpty String senderNumber,

            @Idx(index = "4", type = TEXT)
            @Pkg(label = "[[MakeCall.RecipientNumber.label]]", default_value_type = STRING, description = "[[MakeCall.RecipientNumber.description]]")
            @NotEmpty String recipientNumber,

            @Idx(index = "5", type = TEXT)
            @Pkg(label = "[[MakeCall.CallMessage.label]]", default_value_type = STRING, description = "[[MakeCall.CallMessage.description]]")
            @NotEmpty String  messageBody
    ) {
        String result = "";
        try{
            //Business logic to place voice call
            Twilio.init(authSID.getInsecureString(),authToken.getInsecureString());
            Call call = Call.creator(new PhoneNumber(recipientNumber), new PhoneNumber(senderNumber),
                    new com.twilio.type.Twiml(messageBody)).create();
            result = call.getSid();
        }catch(Exception e){
            throw new BotCommandException(e.getMessage(), e);
        }
        //Return StringValue.
        return new StringValue(result);
    }
}
