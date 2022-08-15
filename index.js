const dotenv = require('dotenv').config()
const express = require('express')
const http = require('http')
const app = express()
const {IncomingWebhook} = require('ms-teams-webhook')

// Initialize IncomingWebhook object with the url of the Teams Webhook information
const url = process.env.MS_TEAMS_WEBHOOK_URL;
const webhook = new IncomingWebhook(url);

// Set view engine and app options
app.set("view engine", "ejs")
app.use(express.json())

// Router entry for webpage
app.get("/", (req, res) => {
    res.render('index', {
        title: 'Teams Webhook'
    })
})

// Router entry for Webhook function
app.post('/json', async (req, res) => {
    // Log the information sent to the webhook from the Webpage
    console.log(req.body.title)
    console.log(req.body.text)
    // Sends information entered into the site to the webhook
    await webhook.send(JSON.stringify({
        "type":"message",
        "attachments":[{
        "contentType":"application/vnd.microsoft.card.adaptive",
        "contentUrl":null,
        "content":{
            "$schema":"http://adaptivecards.io/schemas/adaptive-card.json",
            "type":"AdaptiveCard",
            "version":"1.2",
            "body":[{
                "title": req.body.title,
                "text": req.body.text
            }]
    }}]})).then(() => console.log("Message Sent")).catch(err => console.log(err.reason))

})

// Creates web server for the express app
http.createServer(app).listen(8080, () => {
    console.log('HTTP server up at port 8080.')
    console.log(url)
})