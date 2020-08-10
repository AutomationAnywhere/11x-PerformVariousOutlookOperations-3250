package com.automationanywhere.botcommand.twiliomessenger.commands;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.DataType;
import com.twilio.Twilio;
import com.twilio.rest.api.v2010.account.Message;
import static com.automationanywhere.commandsdk.model.AttributeType.CREDENTIAL;
import static com.automationanywhere.commandsdk.model.AttributeType.TEXT;
import static com.automationanywhere.commandsdk.model.DataType.STRING;
import com.automationanywhere.core.security.SecureString;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPkg adds required information to be displayed on GUI.
@CommandPkg(
//Unique name inside a package and label to display.
        label = "[[SMS.label]]",
        name = "[[SMS.name]]",
        description = "[[SMS.description]]",
        node_label = "[[SMS.node_label]]",
        icon ="twilio.svg",
        //Return type information. return_type ensures only the right kind of variable is provided on the UI.
        return_type = DataType.STRING,
        return_label = "[[SMS.return_label]]",
        return_required = true)

public class SMS {

    //Identify the entry point for the action. Returns a StringValue because the return type is String.
    @Execute
    public StringValue sendSMS(
            //Idx 1 would be displayed first, with a credential box for entering the value.
            @Idx(index = "1", type = CREDENTIAL)
            //UI labels.
            @Pkg(label = "[[SMS.SID.label]]",   default_value_type = STRING, description = "[[SMS.SID.description]]")
            //Ensure that a validation error is thrown when the value is null.
            @NotEmpty SecureString authSID,

            @Idx(index = "2", type = CREDENTIAL)
            @Pkg(label = "[[SMS.AuthToken.label]]",   default_value_type = STRING, description = "[[SMS.AuthToken.description]]")
            @NotEmpty SecureString authToken,

            @Idx(index = "3", type = TEXT)
            @Pkg(label = "[[SMS.SenderNumber.label]]", default_value_type = STRING, description = "[[SMS.SenderNumber.description]]")
            @NotEmpty String senderNumber,

            @Idx(index = "4", type = TEXT)
            @Pkg(label = "[[SMS.RecipientNumber.label]]", default_value_type = STRING, description = "[[SMS.RecipientNumber.description]]")
            @NotEmpty String recipientNumber,

            @Idx(index = "5", type = TEXT)
            @Pkg(label = "[[SMS.MessageBody.label]]", default_value_type = STRING, description = "[[SMS.MessageBody.description]]")
            @NotEmpty String  messageBody
    ) {
        String result = "";
        
        try{
            //Business logic to send out SMS
            Twilio.init(authSID.getInsecureString(),authToken.getInsecureString());
            Message msg = Message.creator(
                    new com.twilio.type.PhoneNumber(recipientNumber),
                    new com.twilio.type.PhoneNumber(senderNumber),
                    messageBody)
                    .create();
            result = msg.getSid();
        }catch(Exception e){
            throw new BotCommandException(e.getMessage(), e);

        }
        //Return StringValue.
        return new StringValue(result);

    }
}
