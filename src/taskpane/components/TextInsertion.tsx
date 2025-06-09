import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles, Text } from "@fluentui/react-components";
/* global HTMLTextAreaElement */
//const [bodyValue, setBodyValue] = useState("Body");



let email = document.getElementById('email-body-sub');
let sub : string;
let text_to_be_ins = "I hope this email finds you well.";

async function fetchUserDetails() {
   const fetch = require('node-fetch');

const url = 'https://catfact.ninja/fact';


try {
	const response = await fetch(url);
	const result = await response.json();
	document.getElementById('email-gen-sub').textContent =  "Email Body : " + "\n" + JSON.stringify(result);
} catch (error) {
	console.error(error);
}
}
Office.onReady(()=> {
  let insert = document.getElementById("email-body")
  
  
  
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
  if (bodyResult.status === Office.AsyncResultStatus.Failed) {
    console.log(`Failed to get body: ${bodyResult.error.message}`);
    return;
  }


  document.getElementById('email-body').innerHTML = "Email Body: " +  bodyResult.value;
  let email_sub = document.getElementById('email-subject');
 Office.context.mailbox.item.subject.getAsync((result) => {
  if (result.status !== Office.AsyncResultStatus.Succeeded) {
    console.error(`Action failed with message ${result.error.message}`);
    return;
  }
  console.log(`Subject: ${result.value}`);
  email_sub.appendChild(document.createTextNode("Subject : "))
   

   email_sub.appendChild(document.createTextNode(result.value));
    email_sub.appendChild(document.createElement("br"))
    email_sub.appendChild(document.createElement("br"))
    

   sub = result.value.toString()
}); 

  email_sub.style.fontFamily = "Verdana, sans-serif";
  
  
 
});
  
})

interface TextInsertionProps {
  insertText: (text: string, insertLocation: Word.InsertLocation.replace | Word.InsertLocation.start | Word.InsertLocation.end | "Replace" | "Start" | "End") => void;
}



const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "10px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "10px",
    maxWidth: "100%",
  },

  email:{
    marginLeft: "20px",
    fontSize: "30px",
  }
  
});

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [text, setText] = useState<string>("Some text.");

  const handleTextInsertion = async () => {
  //  await props.search(text)
  await fetchUserDetails();
    await props.insertText(text, "Replace");
    
  };

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  const styles = useStyles();
  

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="large" label="Enter the Prompt.">
        <Textarea size="large" value={text} onChange={handleTextChange} />
      </Field>
      {/*<Field className={styles.instructions}>Click the button to insert text.</Field>*/}
      <Button appearance="primary" disabled={false} size="large" onClick={handleTextInsertion}>
        Insert 
      </Button>
      
      
    </div>
  );
};

export default TextInsertion;
