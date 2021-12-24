"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.XmlTool = void 0;
const xml2js_1 = __importDefault(require("xml2js"));
const jszip_1 = __importDefault(require("jszip"));
const fs_1 = __importDefault(require("fs"));
class XmlTool {
    constructor() {
        this.getZip = () => {
            return this.zip;
        };
        this.readOriginal = (type) => __awaiter(this, void 0, void 0, function* () {
            let path = __dirname + `/${type}/templates/template.${type}`;
            yield new Promise((resolve, reject) => fs_1.default.readFile(path, (err, data) => __awaiter(this, void 0, void 0, function* () {
                if (err) {
                    console.error(`Template ${path} not read: ${err}`);
                    reject(err);
                    return;
                }
                ;
                return yield this.zip.loadAsync(data).then(d => {
                    resolve(d);
                });
            })));
        });
        this.readXml = (file) => __awaiter(this, void 0, void 0, function* () {
            return this.zip.file(file).async('string').then(data => {
                return this.parser.parseStringPromise(data);
            });
        });
        this.readXml2 = (file) => __awaiter(this, void 0, void 0, function* () {
            return this.zip.file(file).async('arraybuffer').then(data => {
                return this.parser.parseStringPromise(data);
            });
        });
        this.write = (filename, data) => __awaiter(this, void 0, void 0, function* () {
            var xml = this.builder.buildObject(data);
            this.zip.file(filename, Buffer.from(xml), { base64: true });
        });
        this.writeBuffer = (filename, data) => __awaiter(this, void 0, void 0, function* () {
            // var xml = this.builder.buildObject(data);
            this.zip.file(filename, data, { base64: true });
        });
        this.writeStr = (filename, data) => __awaiter(this, void 0, void 0, function* () {
            // var xml = this.builder.buildObject(data);
            this.zip.file(filename, Buffer.from(data), { base64: true });
        });
        this.generateBuffer = () => __awaiter(this, void 0, void 0, function* () {
            return this.zip.generateAsync({ type: 'nodebuffer' });
        });
        this.generateFile = (name, type) => __awaiter(this, void 0, void 0, function* () {
            const buf = yield this.generateBuffer();
            fs_1.default.writeFileSync(name + '.' + type, buf);
            return buf;
        });
        this.zip = new jszip_1.default();
        this.parser = new xml2js_1.default.Parser({ explicitArray: false });
        this.builder = new xml2js_1.default.Builder();
    }
}
exports.XmlTool = XmlTool;
