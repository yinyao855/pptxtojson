declare module 'utif' {
  const UTIF: {
    decode(buff: ArrayBuffer | Uint8Array): unknown[];
    decodeImage(buff: Uint8Array, img: unknown, ifds: unknown[]): void;
    toRGBA8(out: unknown): Uint8Array;
  };
  export default UTIF;
}

declare module 'pngjs' {
  export class PNG {
    width: number;
    height: number;
    data: Buffer | Uint8Array;
    constructor(options: { width: number; height: number });
    static sync: {
      write(png: PNG): Buffer | Uint8Array;
    };
  }
}
