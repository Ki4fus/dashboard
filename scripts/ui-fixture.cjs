const {app} = require('electron');
app.setPath('userData', process.env.NETTING_TEST_PROFILE);
app.whenReady().then(()=>{});
