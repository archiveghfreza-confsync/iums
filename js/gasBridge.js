const CHANNEL = 'MDP_GAS_BRIDGE';

class GasBridge {

  constructor({
    url,
    timeout = 30000
  }) {

    if (!url) {
      throw new Error('GasBridge URL is required.');
    }

    this.url = url;
    this.timeout = timeout;

    this.iframe = null;

    /*
     * Origin واقعی iframe بعد از redirect
     * خود Apps Script اینجا ذخیره می‌شود.
     */
    this.bridgeOrigin = null;

    this.pending = new Map();

    this.readyPromise = null;
    this.readyResolve = null;
    this.readyReject = null;

    this.handleMessage =
      this.handleMessage.bind(this);
  }


  init() {

    if (this.readyPromise) {
      return this.readyPromise;
    }


    this.readyPromise = new Promise(
      (resolve, reject) => {

        this.readyResolve = resolve;
        this.readyReject = reject;

      }
    );


    window.addEventListener(
      'message',
      this.handleMessage
    );


    const iframe =
      document.createElement('iframe');

    iframe.src = this.url;

    iframe.setAttribute(
      'aria-hidden',
      'true'
    );

    /*
     * مخفی، ولی همچنان load می‌شود.
     */
    Object.assign(
      iframe.style,
      {
        position: 'fixed',
        width: '1px',
        height: '1px',
        border: '0',
        opacity: '0',
        pointerEvents: 'none',
        left: '-10000px',
        top: '-10000px'
      }
    );


    this.iframe = iframe;

    document.body.appendChild(iframe);


    /*
     * Timeout برای load اولیه Bridge
     */
    setTimeout(() => {

      if (!this.bridgeOrigin) {

        this.readyReject(
          new Error(
            'Apps Script Bridge initialization timeout.'
          )
        );

      }

    }, 15000);


    return this.readyPromise;
  }


  handleMessage(event) {

    /*
     * بسیار مهم:
     * پیام باید از همان iframe آمده باشد.
     */
    if (
      !this.iframe ||
      event.source !== this.iframe.contentWindow
    ) {
      return;
    }


    const message = event.data;

    if (
      !message ||
      message.channel !== CHANNEL
    ) {
      return;
    }


    /*
     * اولین پیام، origin واقعی
     * googleusercontent را مشخص می‌کند.
     */
    if (message.type === 'READY') {

      this.bridgeOrigin = event.origin;

      if (this.readyResolve) {
        this.readyResolve(this);
      }

      return;
    }


    /*
     * بعد از handshake فقط همان origin
     * پذیرفته می‌شود.
     */
    if (
      this.bridgeOrigin &&
      event.origin !== this.bridgeOrigin
    ) {
      return;
    }


    if (message.type !== 'RESPONSE') {
      return;
    }


    const pending =
      this.pending.get(message.id);


    if (!pending) {
      return;
    }


    clearTimeout(pending.timer);

    this.pending.delete(message.id);


    if (message.ok) {

      pending.resolve(message.data);

    } else {

      const error = new Error(
        message.error?.message ||
        'Apps Script Bridge error.'
      );

      error.code =
        message.error?.code ||
        'BRIDGE_ERROR';

      pending.reject(error);

    }

  }


  async call(
    action,
    payload = {},
    options = {}
  ) {

    await this.init();


    if (!this.bridgeOrigin) {
      throw new Error(
        'Apps Script Bridge is not ready.'
      );
    }


    const id = crypto.randomUUID();

    const timeout =
      options.timeout ||
      this.timeout;


    return new Promise(
      (resolve, reject) => {

        const timer = setTimeout(() => {

          this.pending.delete(id);

          reject(
            new Error(
              `Bridge timeout: ${action}`
            )
          );

        }, timeout);


        this.pending.set(
          id,
          {
            resolve,
            reject,
            timer
          }
        );


        this.iframe.contentWindow.postMessage(
          {
            channel: CHANNEL,
            type: 'CALL',
            id,
            action,
            payload
          },

          /*
           * origin از handshake گرفته شده،
           * بنابراین "*" نداریم.
           */
          this.bridgeOrigin
        );

      }
    );

  }


  destroy() {

    window.removeEventListener(
      'message',
      this.handleMessage
    );


    for (
      const { reject, timer }
      of this.pending.values()
    ) {

      clearTimeout(timer);

      reject(
        new Error(
          'Bridge destroyed.'
        )
      );

    }


    this.pending.clear();


    if (this.iframe) {

      this.iframe.remove();

      this.iframe = null;

    }


    this.bridgeOrigin = null;
    this.readyPromise = null;

  }

}


export default GasBridge;
