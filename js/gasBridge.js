const GAS_BRIDGE_URL =
    'https://script.google.com/macros/s/AKfycbyORoLPh3JxP1KIYy1pbO5tRtt8M7sJAztHyCReHapr2v5F56wJG9JwtBI8LZccjsC0/exec';


const pendingRequests = new Map();

let bridgeFrame = null;

let bridgeReadyPromise = null;


function initGasBridge() {

    if (bridgeReadyPromise) {
        return bridgeReadyPromise;
    }


    bridgeReadyPromise =
        new Promise(function(resolve, reject) {

            bridgeFrame =
                document.createElement('iframe');

            bridgeFrame.src =
                GAS_BRIDGE_URL;

            bridgeFrame.style.position =
                'fixed';

            bridgeFrame.style.width =
                '1px';

            bridgeFrame.style.height =
                '1px';

            bridgeFrame.style.opacity =
                '0';

            bridgeFrame.style.pointerEvents =
                'none';

            bridgeFrame.style.left =
                '-10000px';

            bridgeFrame.style.border =
                '0';


            bridgeFrame.onload =
                function() {

                    resolve();

                };


            bridgeFrame.onerror =
                function() {

                    reject(
                        new Error(
                            'Unable to load Apps Script Bridge.'
                        )
                    );

                };


            document.body.appendChild(
                bridgeFrame
            );

        });


    return bridgeReadyPromise;
}


window.addEventListener(
    'message',
    function(event) {

        if (
            !bridgeFrame ||
            event.source !==
                bridgeFrame.contentWindow
        ) {
            return;
        }


        const response =
            event.data;


        if (
            !response ||
            response.type !==
                'GAS_RESPONSE'
        ) {
            return;
        }


        const pending =
            pendingRequests.get(
                response.id
            );


        if (!pending) {
            return;
        }


        clearTimeout(
            pending.timeout
        );


        pendingRequests.delete(
            response.id
        );


        if (response.success) {

            pending.resolve(
                response.data
            );

        } else {

            pending.reject(
                new Error(
                    response.error ||
                    'Apps Script error.'
                )
            );

        }

    }
);


async function gasCall(
    action,
    payload = {}
) {

    await initGasBridge();


    const id =
        crypto.randomUUID();


    return new Promise(
        function(resolve, reject) {

            const timeout =
                setTimeout(
                    function() {

                        pendingRequests.delete(
                            id
                        );

                        reject(
                            new Error(
                                'Apps Script request timeout.'
                            )
                        );

                    },
                    30000
                );


            pendingRequests.set(
                id,
                {
                    resolve,
                    reject,
                    timeout
                }
            );


            bridgeFrame
                .contentWindow
                .postMessage(
                    {
                        type: 'GAS_CALL',
                        id: id,
                        action: action,
                        payload: payload
                    },
                    '*'
                );

        }
    );

}
