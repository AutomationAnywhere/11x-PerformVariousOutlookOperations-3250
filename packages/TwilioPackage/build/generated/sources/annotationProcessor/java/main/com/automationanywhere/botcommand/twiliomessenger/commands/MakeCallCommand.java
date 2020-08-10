package com.automationanywhere.botcommand.twiliomessenger.commands;

import com.automationanywhere.bot.service.GlobalSessionContext;
import com.automationanywhere.botcommand.BotCommand;
import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.commandsdk.i18n.Messages;
import com.automationanywhere.commandsdk.i18n.MessagesFactory;
import com.automationanywhere.core.security.SecureString;
import java.lang.ClassCastException;
import java.lang.Deprecated;
import java.lang.Object;
import java.lang.String;
import java.lang.Throwable;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

public final class MakeCallCommand implements BotCommand {
  private static final Logger logger = LogManager.getLogger(MakeCallCommand.class);

  private static final Messages MESSAGES_GENERIC = MessagesFactory.getMessages("com.automationanywhere.commandsdk.generic.messages");

  @Deprecated
  public Optional<Value> execute(Map<String, Value> parameters, Map<String, Object> sessionMap) {
    return execute(null, parameters, sessionMap);
  }

  public Optional<Value> execute(GlobalSessionContext globalSessionContext,
      Map<String, Value> parameters, Map<String, Object> sessionMap) {
    logger.traceEntry(() -> parameters != null ? parameters.toString() : null, ()-> sessionMap != null ?sessionMap.toString() : null);
    MakeCall command = new MakeCall();
    HashMap<String, Object> convertedParameters = new HashMap<String, Object>();
    if(parameters.containsKey("authSID") && parameters.get("authSID") != null && parameters.get("authSID").get() != null) {
      convertedParameters.put("authSID", parameters.get("authSID").get());
      if(!(convertedParameters.get("authSID") instanceof SecureString)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","authSID", "SecureString", parameters.get("authSID").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("authSID") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","authSID"));
    }

    if(parameters.containsKey("authToken") && parameters.get("authToken") != null && parameters.get("authToken").get() != null) {
      convertedParameters.put("authToken", parameters.get("authToken").get());
      if(!(convertedParameters.get("authToken") instanceof SecureString)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","authToken", "SecureString", parameters.get("authToken").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("authToken") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","authToken"));
    }

    if(parameters.containsKey("senderNumber") && parameters.get("senderNumber") != null && parameters.get("senderNumber").get() != null) {
      convertedParameters.put("senderNumber", parameters.get("senderNumber").get());
      if(!(convertedParameters.get("senderNumber") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","senderNumber", "String", parameters.get("senderNumber").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("senderNumber") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","senderNumber"));
    }

    if(parameters.containsKey("recipientNumber") && parameters.get("recipientNumber") != null && parameters.get("recipientNumber").get() != null) {
      convertedParameters.put("recipientNumber", parameters.get("recipientNumber").get());
      if(!(convertedParameters.get("recipientNumber") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","recipientNumber", "String", parameters.get("recipientNumber").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("recipientNumber") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","recipientNumber"));
    }

    if(parameters.containsKey("messageBody") && parameters.get("messageBody") != null && parameters.get("messageBody").get() != null) {
      convertedParameters.put("messageBody", parameters.get("messageBody").get());
      if(!(convertedParameters.get("messageBody") instanceof String)) {
        throw new BotCommandException(MESSAGES_GENERIC.getString("generic.UnexpectedTypeReceived","messageBody", "String", parameters.get("messageBody").get().getClass().getSimpleName()));
      }
    }
    if(convertedParameters.get("messageBody") == null) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.validation.notEmpty","messageBody"));
    }

    try {
      Optional<Value> result =  Optional.ofNullable(command.PlaceCall((SecureString)convertedParameters.get("authSID"),(SecureString)convertedParameters.get("authToken"),(String)convertedParameters.get("senderNumber"),(String)convertedParameters.get("recipientNumber"),(String)convertedParameters.get("messageBody")));
      return logger.traceExit(result);
    }
    catch (ClassCastException e) {
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.IllegalParameters","PlaceCall"));
    }
    catch (BotCommandException e) {
      logger.fatal(e.getMessage(),e);
      throw e;
    }
    catch (Throwable e) {
      logger.fatal(e.getMessage(),e);
      throw new BotCommandException(MESSAGES_GENERIC.getString("generic.NotBotCommandException",e.getMessage()),e);
    }
  }
}
