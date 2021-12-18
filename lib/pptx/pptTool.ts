import { XmlTool } from "../xmlTool";
import { ITextModel } from "./models/text.model";

export class PptTool {
    constructor(private xmlTool: XmlTool) { }

    public addSlidePart = async (): Promise<number> => {
        const pptParts = await this.xmlTool.readXml('[Content_Types].xml');
        const slides = pptParts['Types']['Override'].filter(part => {
            return part.$.ContentType === 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml'
        })

        const slidesIds = slides.map(slide => {
            return slide.$.PartName.split('/ppt/slides/slide')[1].split('.xml')[0];
        })

        const id = Math.max(slidesIds) + 1;

        pptParts['Types']['Override'].push(
            {
                '$':
                {
                    ContentType:
                        'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
                    PartName: `/ppt/slides/slide${id}.xml`
                }
            }
        )

        // pptParts['Types']['Override'].push(
        //     {
        //         '$':
        //         {
        //             ContentType:
        //                 'application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml',
        //             PartName: `/ppt/slideMasters/slideMaster${id}.xml`
        //         }
        //     }
        // )

        await this.xmlTool.write(`[Content_Types].xml`, pptParts);
        const relId = await this.addSlidePPTRels(id);
        await this.addSlideToPPT(relId);
        return id;
    }

    private addSlidePPTRels = async (id: number) => {
        const pptRel = await this.xmlTool.readXml('ppt/_rels/presentation.xml.rels');
        const relId = pptRel.Relationships.Relationship.length + 1;
        pptRel.Relationships.Relationship.push({
            '$':
            {
                Id: 'rId' + relId,
                Type:
                    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
                Target: `slides/slide${id}.xml`
            }
        })

        // pptRel.Relationships.Relationship.push({
        //     '$':
        //     {
        //         Id: 'rId' + relId,
        //         Type:
        //             'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
        //         Target: `slideMasters/slideMaster${id + 1}.xml`
        //     }
        // })

        this.xmlTool.write(`ppt/_rels/presentation.xml.rels`, pptRel);
        await this.addSlideRels(id);
        return relId;
    }

    private addSlideRels = async (id: number) => {
        const slideRel = await this.xmlTool.readXml('ppt/slides/_rels/slide1.xml.rels');
        return this.xmlTool.write(`ppt/slides/_rels/slide${id}.xml.rels`, slideRel);

    }

    private addSlideToPPT = async (relId: number) => {
        const ppt = await this.xmlTool.readXml('ppt/presentation.xml');
        // let slideList = ppt['p:presentation']['p:sldIdLst'];
        if (Array.isArray(ppt['p:presentation']['p:sldIdLst']['p:sldId'])) {
            const list = ppt['p:presentation']['p:sldIdLst']['p:sldId'];
            // console.log(list, ppt['p:presentation']['p:sldIdLst']['p:sldId'])
            ppt['p:presentation']['p:sldIdLst']['p:sldId'].push({
                '$':
                    { id: parseInt(list[list.length - 1].$.id, 10) + 1, 'r:id': 'rId' + relId }
            })
        } else {
            ppt['p:presentation']['p:sldIdLst']['p:sldId'] = [
                {
                    '$':

                        ppt['p:presentation']['p:sldIdLst']['p:sldId'].$

                },
                {
                    '$':
                        { id: parseInt(ppt['p:presentation']['p:sldIdLst']['p:sldId'].$.id, 10) + 1, 'r:id': 'rId' + relId }

                }
            ]
        }

        return this.xmlTool.write('ppt/presentation.xml', ppt);
    }

    public removeTemplateSlide = async () => {
        const ppt = await this.xmlTool.readXml('ppt/presentation.xml');
        ppt['p:presentation']['p:sldIdLst']['p:sldId'] = ppt['p:presentation']['p:sldIdLst']['p:sldId'].filter(slide => {
            return slide.$['r:id'] !== 'rId6';
        })

        return this.xmlTool.write('ppt/presentation.xml', ppt);
    }

    public createSlide = async (id: number) => {
        const resSlide = await this.xmlTool.readXml('ppt/slides/slide1.xml');

        await this.xmlTool.write(`ppt/slides/slide${id}.xml`, resSlide);
        // await this.addSlideMaster(id);
        return resSlide;
    }

    private addSlideMaster = async (id: number) => {
        const resMasterSlide = await this.xmlTool.readXml('ppt/slideMasters/slideMaster1.xml');
        await this.xmlTool.write(`ppt/slideMasters/slideMaster${id}.xml`, resMasterSlide);
        const resMasterSlideRel = await this.xmlTool.readXml('ppt/slideMasters/_rels/slideMaster1.xml.rels');
        await this.xmlTool.write(`ppt/slideMasters/_rels/slideMaster${id}.xml.rels`, resMasterSlideRel);
    }


    public addTitle = async (slide, id: number, text: string, opt?: ITextModel) => {
        slide['p:sld']['p:cSld']['p:spTree']['p:sp'][0]['p:txBody']['a:p']['a:r']['a:t'] = text;
        slide['p:sld']['p:cSld']['p:spTree']['p:sp'][0]['p:spPr']['a:xfrm']['a:off'].$.x = opt?.x || slide['p:sld']['p:cSld']['p:spTree']['p:sp'][0]['p:spPr']['a:xfrm']['a:off'].$.x;
        slide['p:sld']['p:cSld']['p:spTree']['p:sp'][0]['p:spPr']['a:xfrm']['a:off'].$.y = opt?.y || slide['p:sld']['p:cSld']['p:spTree']['p:sp'][0]['p:spPr']['a:xfrm']['a:off'].$.y;
        slide['p:sld']['p:cSld']['p:spTree']['p:sp'][0]['p:spPr']['a:xfrm']['a:ext'].$.cx = opt?.cx || slide['p:sld']['p:cSld']['p:spTree']['p:sp'][0]['p:spPr']['a:xfrm']['a:ext'].$.cx;
        slide['p:sld']['p:cSld']['p:spTree']['p:sp'][0]['p:spPr']['a:xfrm']['a:ext'].$.cy = opt?.cy || slide['p:sld']['p:cSld']['p:spTree']['p:sp'][0]['p:spPr']['a:xfrm']['a:ext'].$.cy;
        this.addColorAndSize(slide['p:sld']['p:cSld']['p:spTree']['p:sp'][0], opt);
        return this.xmlTool.write(`ppt/slides/slide${id}.xml`, slide);
    }

    public addSubTitle = async (slide, id: number, text: string, opt?: ITextModel) => {
        slide['p:sld']['p:cSld']['p:spTree']['p:sp'][1]['p:txBody']['a:p']['a:r']['a:t'] = text;
        slide['p:sld']['p:cSld']['p:spTree']['p:sp'][1]['p:spPr']['a:xfrm']['a:off'].$.x = opt?.x || slide['p:sld']['p:cSld']['p:spTree']['p:sp'][1]['p:spPr']['a:xfrm']['a:off'].$.x;
        slide['p:sld']['p:cSld']['p:spTree']['p:sp'][1]['p:spPr']['a:xfrm']['a:off'].$.y = opt?.y || slide['p:sld']['p:cSld']['p:spTree']['p:sp'][1]['p:spPr']['a:xfrm']['a:off'].$.y;
        slide['p:sld']['p:cSld']['p:spTree']['p:sp'][1]['p:spPr']['a:xfrm']['a:ext'].$.cx = opt?.cx || slide['p:sld']['p:cSld']['p:spTree']['p:sp'][1]['p:spPr']['a:xfrm']['a:ext'].$.cx;
        slide['p:sld']['p:cSld']['p:spTree']['p:sp'][1]['p:spPr']['a:xfrm']['a:ext'].$.cy = opt?.cy || slide['p:sld']['p:cSld']['p:spTree']['p:sp'][1]['p:spPr']['a:xfrm']['a:ext'].$.cy;
        this.addColorAndSize(slide['p:sld']['p:cSld']['p:spTree']['p:sp'][1], opt);
        return this.xmlTool.write(`ppt/slides/slide${id}.xml`, slide);
    }

    public addText = async (slide, id: number, text: string, opt?: ITextModel) => {
        const copy = JSON.parse(JSON.stringify(slide['p:sld']['p:cSld']['p:spTree']['p:sp'][1]));
        copy['p:txBody']['a:p']['a:r']['a:t'] = text;
        this.addColorAndSize(copy, opt);
        copy['p:spPr']['a:xfrm']['a:off'].$.x = opt?.x || copy['p:spPr']['a:xfrm']['a:off'].$.x;
        copy['p:spPr']['a:xfrm']['a:off'].$.y = opt?.y || 3190175;
        copy['p:spPr']['a:xfrm']['a:ext'].$.cx = opt?.cx || copy['p:spPr']['a:xfrm']['a:ext'].$.cx;
        copy['p:spPr']['a:xfrm']['a:ext'].$.cy = opt?.cy || copy['p:spPr']['a:xfrm']['a:ext'].$.cy;
        slide['p:sld']['p:cSld']['p:spTree']['p:sp'].push(copy);
        return this.xmlTool.write(`ppt/slides/slide${id}.xml`, slide);
    }

    private addColorAndSize = (data, opt: ITextModel) => {
        if (opt?.color) {
            data['p:txBody']['a:p']['a:r']['a:rPr'] = {
                $: { sz: "1500" },
                'a:solidFill': { 'a:srgbClr': { $: { val: opt.color } } }
            };
        }
        if (opt?.size) {
            if (data['p:txBody']['a:p']['a:r']['a:rPr']) {
                data['p:txBody']['a:p']['a:r']['a:rPr'].$.sz = opt.size.toString();
            } else {
                data['p:txBody']['a:p']['a:r']['a:rPr'] = { $: { sz: opt.size.toString() } };
            }
        }
    }
}