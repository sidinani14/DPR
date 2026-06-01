function testDoPost() {
  var e = {postData: {contents: '{"action":"getLists"}'}};
  var result = doPost(e);
  Logger.log(result.getContent().substring(0, 200));
}
