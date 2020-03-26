const express = require('express');
const cors = require('cors');
const request = require('request');
const httpProxy = require('http-proxy');

const app = express()
const port = process.env.PORT || 3000
const corsOptions = {
    origin: true
}

const proxy = httpProxy.createProxyServer({});

proxy.on("proxyReq", (proxyReq, req, res, options) => {
    const url = "US6-word-view.officeapps.live.com";
    proxyReq.setHeader('origin', `https://${url}`);
    proxyReq.setHeader('authority', url);
    proxyReq.setHeader('referer', `https://${url}`);
});

proxy.on("error", (err, req, res) => {
    console.warn(err);
});

app.use(cors(corsOptions));
app.use(require('morgan')('dev', {
    skip: (req, res) => {
        return (req.url != "/" && req.url != '/do');
    }
}));

app.get('/', (req, res) => res.send('Hello World!'))

app.get('/do', (req, res) => {
    const mainUrl = `https://US6-word-view.officeapps.live.com/wv/wordviewerframe.aspx?embed=1&ui=en%2DUS&rs=en%2DUS&WOPISrc=http%3A%2F%2Fus6%2Dview%2Dwopi%2Ewopi%2Elive%2Enet%3A808%2Foh%2Fwopi%2Ffiles%2F%40%2FwFileId%3FwFileId%3D${encodeURIComponent(req.query.url)}&access_token_ttl=0`
    console.info(`Sending request to URL: ${mainUrl}`);
    request(mainUrl).pipe(res);
});

app.post('/upload')

app.all('*', (req, res) => {
    const proxyUrl = "https://US6-word-view.officeapps.live.com/wv/"

    proxy.web(req, res, {
        target: proxyUrl,
        changeOrigin: true
    })
})

app.listen(port, () => console.log(`Example app listening on port ${port}!`))