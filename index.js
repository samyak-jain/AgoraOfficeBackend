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
    const url = `https://US6-word-view.officeapps.live.com/wv/wordviewerframe.aspx?embed=1&ui=en%2DUS
    &rs=en%2DUS&WOPISrc=http%3A%2F%2Fus6%2Dview%2Dwopi%2Ewopi%2Elive%2Enet%3A808%2Foh%2Fwopi%2Ffiles%
    2F%40%2FwFileId%3FwFileId%3D${req.query.url}&access_token_ttl=0`

    new inliner(url, (error, html) => {
        res.send(html)
    })
})

app.listen(port, () => console.log(`Example app listening on port ${port}!`))