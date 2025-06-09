import * as React from "react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import TextInsertion from "./TextInsertion";
import { makeStyles, Text } from "@fluentui/react-components";
import { Ribbon24Regular, LockOpen24Regular, DesignIdeas24Regular, MailRegular, MailCheckmarkRegular, ClipboardTextEditRegular } from "@fluentui/react-icons";
import { insertText } from "../taskpane";


interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
});



const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  // The list items are static and won't change at runtime,
  // so this should be an ordinary const, not a part of state.
  const listItems: HeroListItem[] = [
    {
      icon: <MailRegular />,
      primaryText: "Reply to mails",
    },
    {
      icon: <MailCheckmarkRegular />,
      primaryText: "Compose a new mail",
    },
    {
      icon: <ClipboardTextEditRegular />,
      primaryText: "Or generate summary for a mail",
    },
  ];

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-new.png" title={props.title} message="Generate Email" />
      {/*<HeroList message="Available Features :" items={listItems} />*/}
      <TextInsertion insertText={insertText} />
    </div>
  );
};

export default App;
