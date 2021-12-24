/// <reference types="node" />
import JSZip from "jszip";
export declare class XmlTool {
    private zip;
    private parser;
    private builder;
    private parts;
    constructor();
    getZip: () => JSZip;
    readOriginal: (type: 'xlsx' | 'pptx') => Promise<void>;
    readXml: (file: string) => Promise<any>;
    readXml2: (file: string) => Promise<any>;
    write: (filename: string, data: any) => Promise<void>;
    writeBuffer: (filename: string, data: Buffer) => Promise<void>;
    writeStr: (filename: string, data: string) => Promise<void>;
    generateBuffer: () => Promise<Buffer>;
    generateFile: (name: string, type: 'xlsx' | 'pptx') => Promise<Buffer>;
}
