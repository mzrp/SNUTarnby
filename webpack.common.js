const path = require('path');
module.exports ={
    entry: './app.js',
    output: {
        filename:'app.js',
        path: path.resolve(__dirname, 'dist'),
    },
    module: {
        rules: [
          {
            test: /\.mp3$/,
            use: [
              {
                loader: 'file-loader',
              },
            ],
          },
        ],
    },
    devServer: {
      port: 9000,
      historyApiFallback: true,
      watchOptions: {
         aggregateTimeout: 300,
         poll: 1000
      },
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Credentials": "true",
        "Access-Control-Allow-Headers": "Content-Type, Authorization, x-id, Content-Length, X-Requested-With",
        "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS"
     }
    },
    node: {
      child_process: 'empty',
      fs: 'empty',
      crypto: 'empty',
      net: 'empty',
      tls: 'empty'
    },
}