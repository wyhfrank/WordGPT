import * as React from "react";
import { DefaultButton, MessageBar, MessageBarType, ProgressIndicator, TextField } from "@fluentui/react";
import { Configuration, OpenAIApi } from "openai";
import Center from "./Center";
import Container from "./Container";
import Login from "./Login";
/* global Word, localStorage, navigator */

export default function App() {
  const defaultPrompt = `Refine the English to make it more legal:`;
  // const defaultPrompt = `Refine the Chinese to make it more legal, do not add additional information:`;
  // const defaultPrompt = `将下面中文改写得更符合合同规范:`;

  const [apiKey, setApiKey] = React.useState<string>("");
  const [prompt, setPrompt] = React.useState<string>(defaultPrompt);
  const [error, setError] = React.useState<string>("");
  const [loading, setLoading] = React.useState<boolean>(false);
  const [generatedText, setGeneratedText] = React.useState<string>("");
  const [selectedText, setselectedText] = React.useState<string>("");

  React.useEffect(() => {
    const key = localStorage.getItem("apiKey");
    if (key) {
      setApiKey(key);
    }
  }, []);

  const openai = React.useMemo(() => {
    return new OpenAIApi(
      new Configuration({
        apiKey,
      })
    );
  }, [apiKey]);

  const saveApiKey = (key) => {
    setApiKey(key);
    localStorage.setItem("apiKey", key);
    setError("");
  };

  const normalCompletion = async (prompt) => {
    let completion = await openai.createCompletion({
      model: "text-davinci-003",
      prompt: prompt,
      max_tokens: 2048,
      temperature: 0,
    });

    let content = completion.data.choices[0].text.trim();
    return content
  };

  const chatCompletion = async (prompt) => {
    let model;
    model = "gpt-3.5-turbo";
    // model = "gpt-4";
    const completion = await openai.createChatCompletion({
      model: model,
      messages: [
        // {"role": "system", "content": "You are a language expert that help refine text without translating it."},
        { role: "user", content: prompt }
      ],
      max_tokens: 2048,
      temperature: 0,
    });
    let content = completion.data.choices[0].message.content.trim();
    return content;
  };

  const onGenerate = async () => {
    await Word.run(async (context) => {
      setGeneratedText("");
      setLoading(true);
      try {
        // Get the current selection
        let selection = context.document.getSelection();
        // Load the selected text and paragraphs
        context.load(selection, "text, paragraphs");
        await context.sync();
        // Access the selected paragraphs
        let paragraphs = selection.paragraphs;
        context.load(paragraphs, "text");
        await context.sync();

        let selectedText = "";
        // Iterate through each paragraph in the selection
        for (let i = 0; i < paragraphs.items.length; i++) {
          let paragraph = paragraphs.items[i];
          selectedText += paragraph.text;

          // Add newline character after each paragraph except the last one
          if (i < paragraphs.items.length - 1) {
            selectedText += "\n";
          }
        }

        // Display the selected text
        console.log("Selected Text: " + selectedText);
        let finalPrompt = prompt + selectedText;
        console.log("Final Prompt:\n" + finalPrompt);

        // return;

        // let generatedText = await normalCompletion(finalPrompt);
        let generatedText = await chatCompletion(finalPrompt);
        // reduce redundant newlines
        generatedText = generatedText.replace(/\n+/g, "\n");

        setselectedText(selectedText);

        console.log("Generated Text:\n" + generatedText);
        setGeneratedText(generatedText);

        // Insert automatically?
        selection.insertText(generatedText, "Replace");
        await context.sync();

      } catch (error) {
        // Handle any errors that occur
        console.log("Error: " + error);
        setError(error);
      };
      setLoading(false);

    });
  };

  const onInsert = async () => {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(generatedText, "Replace");
      await context.sync();
    });
  };

  const onCopy = async () => {
    navigator.clipboard.writeText(generatedText);
  };

  return (
    <Container>
      {apiKey ? (
        <>
          <TextField
            placeholder="Enter prompt template here"
            value={prompt}
            rows={5}
            multiline={true}
            onChange={(_, newValue: string) => setPrompt(newValue || "")}
          ></TextField>
          <Center
            style={{
              marginTop: "10px",
              marginBottom: "10px",
            }}
          >
            <DefaultButton iconProps={{ iconName: "Robot" }} onClick={onGenerate}>
              Generate
            </DefaultButton>
          </Center>
          {loading && <ProgressIndicator label="Generating text..." />}
          {generatedText && (
            <div>
              <Center
                style={{
                  marginTop: "10px",
                  marginBottom: "10px",
                }}
              >
                <DefaultButton iconProps={{ iconName: "Add" }} onClick={onInsert}>
                  Insert text
                </DefaultButton>
                <DefaultButton iconProps={{ iconName: "Copy" }} onClick={onCopy}>
                  Copy text
                </DefaultButton>
              </Center>

              <TextField
                value={selectedText}
                multiline={true}
                rows={5}
              ></TextField>
              <Center>
                ↓↓↓
              </Center>
              <TextField
                value={generatedText}
                multiline={true}
                rows={20}
                onChange={(_, newValue: string) => setGeneratedText(newValue || "")}
              ></TextField>
            </div>
          )}
        </>
      ) : (
        <Login onSave={saveApiKey} />
      )}
      {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
    </Container>
  );
}
