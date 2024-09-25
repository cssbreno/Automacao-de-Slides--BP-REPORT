function sendErrorEmail(emailBody) {
  MailApp.sendEmail({
    to: "breno.soares@sankhya.com.br",
    subject: "Erro na Execução do Script",
    body: emailBody,
  });
}
