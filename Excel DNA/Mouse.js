document.addEventlistener('click', function (event) {
    let elem = event.target;
    let jsonobject =
    {
        key: 'click',
        value: elem.name || elem.id || elem.tagname || "unkown"
    };
    console.log("click")
    window.chrome.webview.postmessage(jsonobject);
});
window.addEventListener('message', event => { window.chrome.webview.postMessage(event.data); })
