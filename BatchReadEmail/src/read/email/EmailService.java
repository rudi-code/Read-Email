package read.email;

import java.io.File;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Properties;

import javax.mail.Address;
import javax.mail.Flags;
import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.MimeBodyPart;

import com.sun.mail.imap.IMAPFolder;

import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.ItemView;

public class EmailService {
	
	public boolean readEmailViaEWS() {
		String username = "rudipurnomo.mail@gmail.com";
		String password = "123jkl.gmail.com";
		
		ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
		ExchangeCredentials credentials = new WebCredentials(username, password);
		try {
			service.setCredentials(credentials);
			service.autodiscoverUrl(username);
			service.setTraceEnabled(true);
			System.out.println("Credential OK");
			ItemView view = new ItemView(5);
			FindItemsResults<Item> findResults = service.findItems(WellKnownFolderName.Inbox, view);
			for(Item item : findResults.getItems()) {
				item.load();
				System.out.println(item.getSubject());
			}
			
		} catch (Exception e) {
			// TODO: handle exception
		}
		return true;
	}
	
	public boolean readEmailViaImaps() {
		try {
        	Properties props = new Properties();
        	props.put("mail.store.protocol","imaps");
            Session session = Session.getDefaultInstance(props, null);
 
            Store store = session.getStore("imaps");
            store.connect("imap.gmail.com", "rudipurnomo.mail", "123jkl.gmail.com");
 
//            Folder inbox = store.getFolder("inbox");
            Folder inbox = (IMAPFolder) store.getFolder("[Gmail]/Important");
            inbox.open(Folder.READ_WRITE);

            int messageCountInbox = inbox.getMessageCount();

            System.out.println("Total Messages Inbox:- " + messageCountInbox);
 
            Message[] messages = inbox.getMessages();
            List<File> attachment = new ArrayList<File>();  
            
            System.out.println("------------------------------");

            for (int i = 0; i < 10; i++) {
            	Message message = messages[messageCountInbox-1-i];
//            	Message message = messages[i];
//            	if(message.getSubject().contains("request slip gaji")) {
            		Multipart multipart = (Multipart) message.getContent();
            		Address[] fromAddress = message.getFrom();
            		String from = fromAddress[0].toString();
            		String sendDate = message.getSentDate().toString();
            		
            		System.out.println("multipart : "+multipart.getCount());
            		System.out.println("subject : "+message.getSubject());
            		System.out.println("from : "+from);
            		System.out.println("date : "+sendDate);
            		
//            		int numberOfParts = multipart.getCount();
//            		for (int partCount = 0; partCount < numberOfParts; partCount++) {
//            			MimeBodyPart part = (MimeBodyPart) multipart.getBodyPart(partCount);
//                        if (part.ATTACHMENT.equalsIgnoreCase(part.getDisposition())) {
//                            // this part is attachment
//                            String fileName = part.getFileName();
//                            part.saveFile("D:\\MyProject\\Attachment" + File.separator + fileName);
//                            System.out.println("file name : "+fileName);
//                            message.setFlag(Flags.Flag.DELETED, true);
//                            
//                        } else {
//                            // this part may be the message content
////                            messageContent = part.getContent().toString();
//                        }
//                    }          		
//            	}
            }
 
            inbox.close(true);
            store.close();
 
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("failed");
        }
		
		return true;
	}

}
