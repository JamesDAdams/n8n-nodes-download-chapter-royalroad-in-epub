import type {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
} from 'n8n-workflow';
import { ApplicationError, NodeConnectionType, NodeOperationError } from 'n8n-workflow';
import { load as loadHtml } from 'cheerio';

// Minimal fetch definition for environments where DOM lib types aren't present
declare const fetch: (
	input: string,
	init?: { headers?: Record<string, string> },
) => Promise<{
	ok: boolean;
	status: number;
	text(): Promise<string>;
}>;

const CSS = `<style>
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
    </style>
`;

function sleep(ms: number): Promise<void> {
	return new Promise((resolve) => (globalThis as any).setTimeout(resolve, ms));
}

async function fetchBook(
	nextUrl: string,
	nbChapters: number,
	startChapter = 1,
): Promise<{ book: string; lastChapterHtml: string; fetched: number }> {
	let book = CSS;
	let chapterCount = 1;
	let lastS = '';

	while (chapterCount <= nbChapters) {
		let res = await fetch(nextUrl, {
			headers: {
				'User-Agent':
					'Mozilla/5.0 (compatible; RoyalRoadNode/1.0; +https://example.com)',
				Accept: 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
				'Accept-Language': 'en-US,en;q=0.9',
			},
		});

		while (res.status === 429) {
			await sleep(200);
			res = await fetch(nextUrl, {
				headers: {
					'User-Agent': 'Mozilla/5.0 (compatible; RoyalRoadNode/1.0)',
					Accept:
						'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
					'Accept-Language': 'en-US,en;q=0.9',
				},
			});
		}

		if (!res.ok) {
			throw new ApplicationError(`HTTP ${res.status} on ${nextUrl}`);
		}

		const html = await res.text();
		const $ = loadHtml(html);

		if (chapterCount >= startChapter) {
			let S = '';
			const chapterTitle = $('h1.font-white').first().text().trim();
			S += `<h1 class="chapter">${chapterTitle}</h1>\n`;

			const chapterContent = $('.chapter-inner').first().html() || '';
			S += chapterContent;

			const authorNotes = $('.portlet-body.author-note');
			if (authorNotes.length === 1) {
				S += '<h3 class="author_note"> Author note </h3>\n';
				S += $(authorNotes[0]).html() || '';
			} else if (authorNotes.length === 2) {
				S += '<h3 class="author_note"> Author note top page </h3>\n';
				S += $(authorNotes[0]).html() || '';
				S += '<h3 class="author_note"> Author note bottom page </h3>\n';
				S += $(authorNotes[1]).html() || '';
			}

			book += S;
			lastS = S;
		}

		const nextRel = $('[rel=next]').first().attr('href');
		if (!nextRel) {
			break;
		}

		nextUrl = /^https?:\/\//i.test(nextRel) ? nextRel : `https://www.royalroad.com${nextRel}`;
		chapterCount += 1;
	}

	return { book, lastChapterHtml: lastS, fetched: chapterCount - 1 };
}

export class RoyalRoadNode implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'RoyalRoad: Fetch Chapters',
		name: 'royalRoadNode',
		group: ['transform'],
		version: 1,
		description: 'Fetches sequential RoyalRoad chapters and outputs combined HTML',
		defaults: {
			name: 'RoyalRoad: Fetch Chapters',
		},
		inputs: [NodeConnectionType.Main],
		outputs: [NodeConnectionType.Main],
		usableAsTool: true,
		properties: [
			{
				displayName: 'Start Chapter URL',
				name: 'url',
				type: 'string',
				default: '',
				placeholder: 'https://www.royalroad.com/fiction/.../chapter/...',
				description: 'URL of the first chapter to start fetching from',
				required: true,
			},
			{
				displayName: 'Number of Chapters',
				name: 'chapters',
				type: 'number',
				default: 1,
				description: 'How many chapters to fetch starting from the given URL',
				typeOptions: {
					minValue: 1,
					maxValue: 1000,
				},
				required: true,
			},
			{
				displayName: 'Start Chapter Index',
				name: 'startChapter',
				type: 'number',
				default: 1,
				description: 'Start including content from this chapter index (1-based)',
				typeOptions: {
					minValue: 1,
				},
			},
			{
				displayName: 'Output',
				name: 'output',
				type: 'options',
				options: [
					{ name: 'Binary File (HTML)', value: 'binary' },
					{ name: 'String (HTML) in JSON', value: 'json' },
				],
				default: 'binary',
				description: 'How to output the fetched content',
			},
			{
				displayName: 'File Name',
				name: 'fileName',
				type: 'string',
				default: 'royalroad.html',
				description: 'Name of the generated HTML file when output is binary',
				displayOptions: {
					show: {
						output: ['binary'],
					},
				},
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();

		for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
			try {
				const url = this.getNodeParameter('url', itemIndex, '') as string;
				const chapters = this.getNodeParameter('chapters', itemIndex, 1) as number;
				const startChapter = this.getNodeParameter('startChapter', itemIndex, 1) as number;
				const output = this.getNodeParameter('output', itemIndex, 'binary') as 'binary' | 'json';
				const fileName = this.getNodeParameter('fileName', itemIndex, 'royalroad.html') as string;

				if (!url) {
					throw new NodeOperationError(this.getNode(), 'URL is required', { itemIndex });
				}

				const { book, lastChapterHtml, fetched } = await fetchBook(url, chapters, startChapter);

				if (output === 'binary') {
					const binBuffer = (globalThis as any).Buffer.from(book, 'utf8');
					const binary = await this.helpers.prepareBinaryData(
						binBuffer,
						fileName,
						'text/html; charset=utf-8',
					);
					items[itemIndex].binary = items[itemIndex].binary ?? {};
					(items[itemIndex].binary as any).data = binary;
					items[itemIndex].json = {
						url,
						requestedChapters: chapters,
						startChapter,
						fetchedChapters: fetched,
						lastChapterPreview: lastChapterHtml ? lastChapterHtml.slice(0, 200) : '',
					};
				} else {
					items[itemIndex].json = {
						url,
						requestedChapters: chapters,
						startChapter,
						fetchedChapters: fetched,
						book,
						lastChapterHtml,
					};
				}
			} catch (error) {
				if (this.continueOnFail()) {
					items[itemIndex] = {
						json: items[itemIndex]?.json ?? {},
						error,
						pairedItem: itemIndex,
					} as unknown as INodeExecutionData;
					continue;
				}
				if ((error as any)?.context) {
					(error as any).context.itemIndex = itemIndex;
					throw error;
				}
				throw new NodeOperationError(this.getNode(), error as Error, { itemIndex });
			}
		}

		return [items];
	}
}
