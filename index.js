const express = require('express')
const inliner = require('inliner')
const cors = require('cors')
const app = express()
const port = process.env.PORT || 3000
const corsOptions = {
    origin: '*'
}

app.use(cors(corsOptions));

app.get('/', (req, res) => res.send('Hello World!'))

app.get('/do', (req, res) => {
    new inliner('https://view.officeapps.live.com/op/embed.aspx?src=' + req.query.url, (error, html) => {
        res.send(html)
    })
})

app.listen(port, () => console.log(`Example app listening on port ${port}!`))