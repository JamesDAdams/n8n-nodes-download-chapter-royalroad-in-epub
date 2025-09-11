/// <reference types="node" />

import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { NodeConnectionType, NodeOperationError } from 'n8n-workflow';
import { load as loadHtml } from 'cheerio';
import { deflateRawSync } from 'zlib';

const DEFAULT_CSS = `
.chapter-content table {
  background: #004b7a;
  margin: 10px auto;
  width: 90%;
  border: none;
  box-shadow: 1px 1px 1px rgba(0,0,0,.75);
  border-collapse: separate;
  border-spacing: 2px;
}
.chapter-content table td {
  color: #ccc;
}
`;

function uuidv4(): string {
	return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
		const r = (Math.random() * 16) | 0;
		const v = c === 'x' ? r : (r & 0x3) | 0x8;
		return v.toString(16);
	});
}

function slugify(str: string): string {
	return String(str)
		.toLowerCase()
		.replace(/[^a-z0-9]+/g, '-')
		.replace(/^-+|-+$/g, '')
		.substring(0, 60) || 'chapter';
}

function escapeXml(unsafe: string): string {
	return String(unsafe)
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&apos;');
}

// Minimal CRC32 implementation
const CRC_TABLE = (() => {
	const table = new Uint32Array(256);
	for (let n = 0; n < 256; n++) {
		let c = n;
		for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
		table[n] = c >>> 0;
	}
	return table;
})();
function crc32(buf: Buffer): number {
	let c = 0 ^ -1;
	for (let i = 0; i < buf.length; i++) c = (c >>> 8) ^ CRC_TABLE[(c ^ buf[i]) & 0xff];
	return (c ^ -1) >>> 0;
}

function msDosDateTime(d = new Date()): { time: number; date: number } {
	const year = d.getUTCFullYear();
	const month = d.getUTCMonth() + 1;
	const day = d.getUTCDate();
	const hours = d.getUTCHours();
	const minutes = d.getUTCMinutes();
	const seconds = Math.floor(d.getUTCSeconds() / 2);
	const time = (hours << 11) | (minutes << 5) | seconds;
	const date = ((year - 1980) << 9) | (month << 5) | day;
	return { time, date };
}

class ZipBuilder {
	private files: Array<{
		name: string;
		crc: number;
		comp: Buffer;
		uncomp: number;
		method: number;
		time: number;
		date: number;
		offset: number;
	}> = [];
	private chunks: Buffer[] = [];
	private offset = 0;
	private readonly UTF8_FLAG = 0x0800;
	addFile(name: string, data: string | Buffer, opts?: { compress?: boolean; date?: Date }) {
		const buf = Buffer.isBuffer(data) ? data : Buffer.from(data, 'utf8');
		const { time, date } = msDosDateTime(opts?.date);
		const crc = crc32(buf);
		const compress = opts?.compress !== false; // default true
		const method = compress ? 8 : 0; // 8 = deflate, 0 = store
		const comp = compress ? deflateRawSync(buf) : buf;
		const nameBuf = Buffer.from(name, 'utf8');
		const localHeader = Buffer.alloc(30 + nameBuf.length);
		let p = 0;
		localHeader.writeUInt32LE(0x04034b50, p); p += 4; // local header sig
		localHeader.writeUInt16LE(20, p); p += 2; // version needed
		localHeader.writeUInt16LE(this.UTF8_FLAG, p); p += 2; // flags (UTF-8)
		localHeader.writeUInt16LE(method, p); p += 2; // method
		localHeader.writeUInt16LE(time, p); p += 2; // time
		localHeader.writeUInt16LE(date, p); p += 2; // date
		localHeader.writeUInt32LE(crc >>> 0, p); p += 4; // crc
		localHeader.writeUInt32LE(comp.length >>> 0, p); p += 4; // comp size
		localHeader.writeUInt32LE(buf.length >>> 0, p); p += 4; // uncomp size
		localHeader.writeUInt16LE(nameBuf.length, p); p += 2; // fname len
		localHeader.writeUInt16LE(0, p); p += 2; // extra len
		nameBuf.copy(localHeader, p);

		const offset = this.offset;
		this.chunks.push(localHeader, comp);
		this.offset += localHeader.length + comp.length;

		this.files.push({ name, crc, comp, uncomp: buf.length, method, time, date, offset });
	}
	build(): Buffer {
		const central: Buffer[] = [];
		let centralSize = 0;
		for (const f of this.files) {
			const nameBuf = Buffer.from(f.name, 'utf8');
			const hdr = Buffer.alloc(46 + nameBuf.length);
			let p = 0;
			hdr.writeUInt32LE(0x02014b50, p); p += 4; // central header sig
			hdr.writeUInt16LE(0x031E, p); p += 2; // version made by (arbitrary)
			hdr.writeUInt16LE(20, p); p += 2; // version needed
			hdr.writeUInt16LE(0x0800, p); p += 2; // flags UTF-8
			hdr.writeUInt16LE(f.method, p); p += 2; // method
			hdr.writeUInt16LE(f.time, p); p += 2; // time
			hdr.writeUInt16LE(f.date, p); p += 2; // date
			hdr.writeUInt32LE(f.crc >>> 0, p); p += 4; // crc
			hdr.writeUInt32LE(f.comp.length >>> 0, p); p += 4; // comp size
			hdr.writeUInt32LE(f.uncomp >>> 0, p); p += 4; // uncomp size
			hdr.writeUInt16LE(nameBuf.length, p); p += 2; // name len
			hdr.writeUInt16LE(0, p); p += 2; // extra len
			hdr.writeUInt16LE(0, p); p += 2; // comment len
			hdr.writeUInt16LE(0, p); p += 2; // disk start
			hdr.writeUInt16LE(0, p); p += 2; // int attrs
			hdr.writeUInt32LE(0, p); p += 4; // ext attrs
			hdr.writeUInt32LE(f.offset >>> 0, p); p += 4; // rel offset
			nameBuf.copy(hdr, p);
			central.push(hdr);
			centralSize += hdr.length;
		}
		const centralStart = this.offset;
		this.chunks.push(...central);
		this.offset += centralSize;

		const eocd = Buffer.alloc(22);
		let p = 0;
		eocd.writeUInt32LE(0x06054b50, p); p += 4; // EOCD sig
		eocd.writeUInt16LE(0, p); p += 2; // disk no
		eocd.writeUInt16LE(0, p); p += 2; // disk start
		eocd.writeUInt16LE(this.files.length, p); p += 2; // total entries disk
		eocd.writeUInt16LE(this.files.length, p); p += 2; // total entries
		eocd.writeUInt32LE(centralSize >>> 0, p); p += 4; // central size
		eocd.writeUInt32LE(centralStart >>> 0, p); p += 4; // central offset
		eocd.writeUInt16LE(0, p); p += 2; // comment len
		this.chunks.push(eocd);
		this.offset += eocd.length;

		return Buffer.concat(this.chunks);
	}
}

function buildXhtml(title: string, bodyHtml: string): string {
	return `<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <title>${escapeXml(title)}</title>
  <link rel="stylesheet" type="text/css" href="style.css" />
</head>
<body>
${bodyHtml}
</body>
</html>`;
}

function makeOpf(args: {
	uuid: string;
	title: string;
	author: string;
	language: string;
	chapters: Array<{ title: string; filename: string }>;
}): string {
	const manifestItems = [
		'<item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>',
		'<item id="style" href="style.css" media-type="text/css"/>',
	]
		.concat(
			args.chapters.map(
				(c, i) => `<item id="chapter-${i + 1}" href="${c.filename}" media-type="application/xhtml+xml"/>`,
			),
		)
		.join('\n    ');

	const spineItems = args.chapters
		.map((c, i) => `<itemref idref="chapter-${i + 1}" />`)
		.join('\n    ');

	return `<?xml version=\"1.0\" encoding=\"utf-8\"?>
<package xmlns=\"http://www.idpf.org/2007/opf\" unique-identifier=\"BookId\" version=\"2.0\">
  <metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:opf=\"http://www.idpf.org/2007/opf\">
    <dc:title>${escapeXml(args.title)}</dc:title>
    <dc:creator opf:role=\"aut\">${escapeXml(args.author)}</dc:creator>
    <dc:language>${escapeXml(args.language)}</dc:language>
    <dc:identifier id=\"BookId\">urn:uuid:${args.uuid}</dc:identifier>
  </metadata>
  <manifest>
    ${manifestItems}
  </manifest>
  <spine toc=\"ncx\">
    ${spineItems}
  </spine>
</package>`;
}

function makeNcx(args: {
	uuid: string;
	title: string;
	author: string;
	chapters: Array<{ title: string; filename: string }>;
}): string {
	const navPoints = args.chapters
		.map(
			(c, i) => `    <navPoint id=\"navPoint-${i + 1}\" playOrder=\"${i + 1}\">
      <navLabel><text>${escapeXml(c.title || `Chapter ${i + 1}`)}</text></navLabel>
      <content src=\"${c.filename}\" />
    </navPoint>`,
		)
		.join('\n');

	return `<?xml version=\"1.0\" encoding=\"utf-8\"?>
<!DOCTYPE ncx PUBLIC \"-//NISO//DTD ncx 2005-1//EN\" \"http://www.daisy.org/z3986/2005/ncx-2005-1.dtd\">
<ncx xmlns=\"http://www.daisy.org/z3986/2005/ncx/\" version=\"2005-1\">
  <head>
    <meta name=\"dtb:uid\" content=\"urn:uuid:${args.uuid}\"/>
    <meta name=\"dtb:depth\" content=\"1\"/>
    <meta name=\"dtb:totalPageCount\" content=\"0\"/>
    <meta name=\"dtb:maxPageNumber\" content=\"0\"/>
  </head>
  <docTitle><text>${escapeXml(args.title)}</text></docTitle>
  <docAuthor><text>${escapeXml(args.author)}</text></docAuthor>
  <navMap>
${navPoints}
  </navMap>
</ncx>`;
}

function extractChaptersAndCss(html: string): { chapters: Array<{ title: string; html: string }>; css: string } {
	const $probe = loadHtml(html);
	let cssCollected = '';
	$probe('style').each((_, el) => {
		const css = $probe(el).html();
		if (css) cssCollected += css + '\n';
	});
	cssCollected = (cssCollected || DEFAULT_CSS).trim();

	const $ = loadHtml(`<div id=\"root\">${html}</div>`);
	const chapters: Array<{ title: string; html: string }> = [];
	$('h1.chapter').each((i, el) => {
		const title = $(el).text().trim() || `Chapter ${i + 1}`;
		const titleHtml = $.html(el);
		const siblings = $(el).nextUntil('h1.chapter');
		const bodyHtml = siblings
			.map((_, s) => $.html(s))
			.get()
			.join('');
		chapters.push({ title, html: `${titleHtml}\n${bodyHtml}` });
	});
	if (chapters.length === 0) {
		chapters.push({ title: 'Chapter 1', html });
	}
	return { chapters, css: cssCollected };
}

function writeEpubBufferFromParts(
	chaptersIn: Array<{ title: string; html: string }>,
	css: string,
	opts: { title?: string; author?: string; language?: string },
): { buffer: Buffer; uuid: string; title: string; chapterCount: number } {
	const language = opts.language || 'en';
	const author = opts.author || 'Unknown';
	const uuid = uuidv4();
	const inferredTitle = opts.title || chaptersIn[0]?.title || 'Book';

	const chapterFiles = chaptersIn.map((c, i) => {
		const filename = `chapter-${i + 1}-${slugify(c.title)}.xhtml`;
		const xhtml = buildXhtml(c.title, c.html);
		return { title: c.title, filename, xhtml };
	});

	const opf = makeOpf({ uuid, title: inferredTitle, author, language, chapters: chapterFiles });
	const ncx = makeNcx({ uuid, title: inferredTitle, author, chapters: chapterFiles });
	const containerXml = `<?xml version=\"1.0\"?>
<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">
  <rootfiles>
    <rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/>
  </rootfiles>
</container>`;

	const zip = new ZipBuilder();
	// EPUB spec requires 'mimetype' as first entry and STORED (no compression)
	zip.addFile('mimetype', 'application/epub+zip', { compress: false });
	zip.addFile('META-INF/container.xml', containerXml);
	zip.addFile('OEBPS/style.css', css);
	zip.addFile('OEBPS/content.opf', opf);
	zip.addFile('OEBPS/toc.ncx', ncx);
	for (const ch of chapterFiles) {
		zip.addFile(`OEBPS/${ch.filename}`, ch.xhtml);
	}
	const buffer = zip.build();
	return { buffer, uuid, title: inferredTitle, chapterCount: chaptersIn.length };
}

export class HtmlToEpubNode implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'HTML → EPUB',
		name: 'htmlToEpubNode',
		icon: { light: 'file:html-to-svg.svg', dark: 'file:html-to-svg.svg' },
		group: ['transform'],
		version: 1,
		description: 'Convert HTML (binary or string) to EPUB file',
		defaults: { name: 'HTML → EPUB' },
		inputs: [NodeConnectionType.Main],
		outputs: [NodeConnectionType.Main],
		usableAsTool: true,
		properties: [
			{
				displayName: 'Input Mode',
				name: 'inputMode',
				type: 'options',
				options: [
					{ name: 'Binary (HTML)', value: 'binary' },
					{ name: 'String (HTML)', value: 'string' },
				],
				default: 'binary',
				description: 'Source of the input HTML',
			},
			{
				displayName: 'Binary Property Name',
				name: 'binaryPropertyName',
				type: 'string',
				default: 'data',
				description: 'Binary property containing the HTML',
				displayOptions: { show: { inputMode: ['binary'] } },
			},
			{
				displayName: 'JSON Property Name with HTML',
				name: 'stringPropertyName',
				type: 'string',
				default: 'book',
				description: 'JSON key containing the raw HTML',
				displayOptions: { show: { inputMode: ['string'] } },
			},
			{
				displayName: 'Title',
				name: 'title',
				type: 'string',
				default: '',
				description: 'Title to set in the EPUB (optional)',
			},
			{
				displayName: 'Author',
				name: 'author',
				type: 'string',
				default: 'Unknown',
				description: 'Author of the EPUB',
			},
			{
				displayName: 'Language',
				name: 'language',
				type: 'string',
				default: 'en',
				description: 'Language of the book',
			},
			{
				displayName: 'EPUB File Name',
				name: 'fileName',
				type: 'string',
				default: 'book.epub',
				description: 'Name of the output EPUB file',
			},
			{
				displayName: 'Output Binary Property Name',
				name: 'outputBinaryPropertyName',
				type: 'string',
				default: 'data',
				description: 'Name of the binary property to write the EPUB to',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();

		for (let i = 0; i < items.length; i++) {
			try {
				const inputMode = this.getNodeParameter('inputMode', i, 'binary') as 'binary' | 'string';
				const fileName = this.getNodeParameter('fileName', i, 'book.epub') as string;
				const outputBinaryPropertyName = this.getNodeParameter('outputBinaryPropertyName', i, 'data') as string;
				const title = (this.getNodeParameter('title', i, '') as string) || undefined;
				const author = this.getNodeParameter('author', i, 'Unknown') as string;
				const language = this.getNodeParameter('language', i, 'en') as string;

				let html = '';
				if (inputMode === 'binary') {
					const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i, 'data') as string;
					const buffer = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);
					html = buffer.toString('utf8');
				} else {
					const stringPropertyName = this.getNodeParameter('stringPropertyName', i, 'book') as string;
					const val = items[i]?.json?.[stringPropertyName];
					if (typeof val !== 'string' || !val) {
						throw new NodeOperationError(this.getNode(), `The JSON property '${stringPropertyName}' must contain a non-empty HTML string.`, { itemIndex: i });
					}
					html = val;
				}

				const { chapters, css } = extractChaptersAndCss(html);
				const { buffer, uuid, title: inferredTitle, chapterCount } = writeEpubBufferFromParts(chapters, css, { title, author, language });

				const binary = await this.helpers.prepareBinaryData(buffer, fileName, 'application/epub+zip');
				items[i].binary = items[i].binary ?? {};
				(items[i].binary as any)[outputBinaryPropertyName] = binary;
				items[i].json = {
					...items[i].json,
					meta: {
						uuid,
						title: inferredTitle,
						author,
						language,
						chapters: chapterCount,
					},
				};
			} catch (error) {
				if (this.continueOnFail()) {
					items[i] = {
						json: items[i]?.json ?? {},
						error,
						pairedItem: i,
					} as unknown as INodeExecutionData;
					continue;
				}
				if ((error as any)?.context) {
					(error as any).context.itemIndex = i;
					throw error;
				}
				throw new NodeOperationError(this.getNode(), error as Error, { itemIndex: i });
			}
		}

		return [items];
	}
}
