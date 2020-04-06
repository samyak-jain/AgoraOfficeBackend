const app = require("./index");
const os = require("os")
const cluster = require("cluster")

const port = process.env.PORT || 3000
 
const clusterWorkerSize = os.cpus().length
 
if (clusterWorkerSize > 1) {
  if (cluster.isMaster) {
    for (let i=0; i < clusterWorkerSize; i++) {
      cluster.fork()
    }
 
    cluster.on("exit", function(worker) {
      console.log("Worker", worker.id, " has exitted.")
    })
  } else {
    app.listen(port, () => console.log(`Example app listening on port ${port}!`));
  }
} else {
    app.listen(port, () => console.log(`Example app listening on port ${port}!`));
}