/************** CREDENTIALS **************/

function getTrelloCredentials() {
  const props = PropertiesService.getScriptProperties();
  return {
    key: props.getProperty("TRELLO_API_KEY"),
    token: props.getProperty("TRELLO_TOKEN")
  };
}
