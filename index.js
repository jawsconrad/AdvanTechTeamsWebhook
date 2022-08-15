const dotenv = require('dotenv').config()
const express = require('express')
const http = require('http')
const app = express()
const {IncomingWebhook} = require('ms-teams-webhook')

const url = process.env.MS_TEAMS_WEBHOOK_URL;
const webhook = new IncomingWebhook(url);

app.set("view engine", "ejs")
app.use(express.json())

app.get("/", (req, res) => {
    res.render('index', {
        title: 'Teams Webhook'
    })
})
app.post('/index', async (req, res) => {
    console.log(req)
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
    }}]}))
})
http.createServer(app)
.listen(8080, () => {
    console.log('HTTP server up at port 8080.')
    console.log(url)
})