import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import * as lancedb from "@lancedb/lancedb";
import { runJxa } from "run-jxa";
import path from "node:path";
import os from "node:os";
import TurndownService from "turndown";
import fs from "node:fs/promises";
import {
  EmbeddingFunction,
  LanceSchema,
  register,
} from "@lancedb/lancedb/embedding";
import { type Float, Float32, Utf8 } from "apache-arrow";
import { pipeline } from "@huggingface/transformers";

const { turndown } = new TurndownService();
const db = await lancedb.connect(
  path.join(os.homedir(), ".mcp-apple-notes", "data")
);
const extractor = await pipeline(
  "feature-extraction",
  "Xenova/all-MiniLM-L6-v2"
);

@register("openai")
export class OnDeviceEmbeddingFunction extends EmbeddingFunction<string> {
  toJSON(): object {
    return {};
  }
  ndims() {
    return 384;
  }
  embeddingDataType(): Float {
    return new Float32();
  }
  async computeQueryEmbeddings(data: string) {
    const output = await extractor(data, { pooling: "mean" });
    return output.data as number[];
  }
  async computeSourceEmbeddings(data: string[]) {
    return await Promise.all(
      data.map(async (item) => {
        const output = await extractor(item, { pooling: "mean" });

        return output.data as number[];
      })
    );
  }
}

const func = new OnDeviceEmbeddingFunction();

const notesTableSchema = LanceSchema({
  title: func.sourceField(new Utf8()),
  content: func.sourceField(new Utf8()),
  creation_date: func.sourceField(new Utf8()),
  modification_date: func.sourceField(new Utf8()),
  vector: func.vectorField(),
});

const QueryNotesSchema = z.object({
  query: z.string(),
});

const GetNoteSchema = z.object({
  title: z.string(),
});

const server = new Server(
  {
    name: "my-apple-notes-mcp",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: "list-notes",
        description: "Lists just the titles of all my Apple Notes",
        inputSchema: {
          type: "object",
          properties: {},
          required: [],
        },
      },
      {
        name: "index-notes",
        description:
          "Index all my Apple Notes for Semantic Search. Please tell the user that the sync takes couple of seconds up to couple of minutes depending on how many notes you have.",
        inputSchema: {
          type: "object",
          properties: {},
          required: [],
        },
      },
      {
        name: "get-note",
        description: "Get a note full content and details by title",
        inputSchema: {
          type: "object",
          properties: {
            title: z.string(),
          },
          required: ["title"],
        },
      },
      {
        name: "search-notes",
        description: "Search for notes by title or content",
        inputSchema: {
          type: "object",
          properties: {
            query: { type: "string" },
          },
          required: ["query"],
        },
      },
      {
        name: "create-note",
        description:
          "Create a new Apple Note with specified title and content. Must be in HTML format WITHOUT newlines",
        inputSchema: {
          type: "object",
          properties: {
            title: { type: "string" },
            content: { type: "string" },
          },
          required: ["title", "content"],
        },
      },
    ],
  };
});

// Define the data directory
const DATA_DIR = path.join(os.homedir(), ".mcp-apple-notes");
const STATE_FILE = path.join(DATA_DIR, "indexing-state.json");

// Create the data directory if it doesn't exist
(async () => {
  try {
    await fs.mkdir(DATA_DIR, { recursive: true });
  } catch (error) {
    console.error(`Error creating data directory: ${error.message}`);
  }
})();

// Define the indexing state interface
interface IndexingState {
  inProgress: boolean;
  totalNotes: number;
  processedNotes: number;
  lastProcessedIndex: number;
  startTime: number;
  lastUpdateTime: number;
  errors: string[];
  batchSize: number;
}

// Initialize default state
const defaultState: IndexingState = {
  inProgress: false,
  totalNotes: 0,
  processedNotes: 0,
  lastProcessedIndex: -1,
  startTime: 0,
  lastUpdateTime: 0,
  errors: [],
  batchSize: 5,
};

// Function to read the current indexing state
async function getIndexingState(): Promise<IndexingState> {
  try {
    const data = await fs.readFile(STATE_FILE, 'utf-8');
    return JSON.parse(data) as IndexingState;
  } catch (error) {
    // If the file doesn't exist or can't be read, return the default state
    return { ...defaultState };
  }
}

// Function to save the current indexing state
async function saveIndexingState(state: IndexingState): Promise<void> {
  try {
    await fs.writeFile(STATE_FILE, JSON.stringify(state, null, 2), 'utf-8');
  } catch (error) {
    console.error(`Error saving indexing state: ${error.message}`);
  }
}

// Function to get the status of the indexing process
async function getIndexingStatus(): Promise<{
  inProgress: boolean;
  progress: number;
  processedNotes: number;
  totalNotes: number;
  elapsedTime: number;
  errors: string[];
}> {
  const state = await getIndexingState();
  const progress = state.totalNotes > 0 
    ? Math.round((state.processedNotes / state.totalNotes) * 100) 
    : 0;
  
  return {
    inProgress: state.inProgress,
    progress,
    processedNotes: state.processedNotes,
    totalNotes: state.totalNotes,
    elapsedTime: state.inProgress ? Date.now() - state.startTime : 0,
    errors: state.errors,
  };
}

const getNotes = async () => {
  console.error("Executing JXA to get notes list...");
  try {
    const notes = await runJxa(`
      try {
        const app = Application('Notes');
        app.includeStandardAdditions = true;
        
        const allNotes = Array.from(app.notes());
        const titles = allNotes.map(note => note.properties().name);
        
        return titles;
      } catch (error) {
        return JSON.stringify({ error: error.toString() });
      }
    `);
    
    // Check if we got an error object
    if (typeof notes === 'string' && notes.includes('"error":')) {
      const errorObj = JSON.parse(notes);
      console.error(`JXA error: ${errorObj.error}`);
      return [];
    }
    
    console.error(`JXA returned ${Array.isArray(notes) ? notes.length : 0} notes`);
    return notes as string[];
  } catch (error) {
    console.error(`Error in getNotes: ${error.message}`);
    return [];
  }
};

const getNoteDetailsByTitle = async (title: string) => {
  const note = await runJxa(
    `const app = Application('Notes');
    const title = "${title}"
    
    try {
        const note = app.notes.whose({name: title})[0];
        
        const noteInfo = {
            title: note.name(),
            content: note.body(),
            creation_date: note.creationDate().toLocaleString(),
            modification_date: note.modificationDate().toLocaleString()
        };
        
        return JSON.stringify(noteInfo);
    } catch (error) {
        return "{}";
    }`
  );

  return JSON.parse(note as string) as {
    title: string;
    content: string;
    creation_date: string;
    modification_date: string;
  };
};

// Background indexing function that processes notes in batches
async function backgroundIndexNotes(notesTable: lancedb.Table): Promise<void> {
  // Get the current state or initialize a new one
  let state = await getIndexingState();
  
  // If indexing is already in progress, don't start again
  if (state.inProgress) {
    console.error("Indexing is already in progress");
    return;
  }
  
  // Initialize the state for a new indexing run
  const allNotes = await getNotes();
  state = {
    ...defaultState,
    inProgress: true,
    totalNotes: allNotes.length,
    startTime: Date.now(),
    lastUpdateTime: Date.now(),
  };
  
  // Save the initial state
  await saveIndexingState(state);
  
  // Process notes in batches
  console.error(`Starting background indexing of ${allNotes.length} notes in batches of ${state.batchSize}`);
  
  try {
    for (let i = 0; i < allNotes.length; i += state.batchSize) {
      // Update the state
      state.lastProcessedIndex = i;
      state.lastUpdateTime = Date.now();
      await saveIndexingState(state);
      
      // Get the current batch
      const batch = allNotes.slice(i, i + state.batchSize);
      console.error(`Processing batch ${Math.floor(i/state.batchSize) + 1} of ${Math.ceil(allNotes.length/state.batchSize)}`);
      
      // Process the batch
      const batchDetails = await Promise.all(
        batch.map(async (noteTitle) => {
          try {
            console.error(`Getting details for note: "${noteTitle}"`);
            const details = await getNoteDetailsByTitle(noteTitle);
            return details;
          } catch (error) {
            const errorMsg = `Error getting note details for ${noteTitle}: ${error.message}`;
            console.error(errorMsg);
            state.errors.push(errorMsg);
            return null;
          }
        })
      );
      
      // Filter out null results and process the notes
      const validDetails = batchDetails.filter(Boolean);
      console.error(`Got ${validDetails.length} valid notes in this batch`);
      
      if (validDetails.length > 0) {
        // Convert HTML to Markdown and prepare for database
        const batchChunks = validDetails.map((note, index) => {
          try {
            // TypeScript non-null assertion to handle the null check we already did with filter(Boolean)
            return {
              id: (i + index).toString(),
              title: note!.title || "Untitled",
              content: note!.content ? turndown(note!.content) : "",
              creation_date: note!.creation_date || new Date().toISOString(),
              modification_date: note!.modification_date || new Date().toISOString(),
            };
          } catch (error) {
            const errorMsg = `Error processing note ${note!.title}: ${error.message}`;
            console.error(errorMsg);
            state.errors.push(errorMsg);
            return {
              id: (i + index).toString(),
              title: note!.title || "Untitled",
              content: note!.content || "",
              creation_date: note!.creation_date || new Date().toISOString(),
              modification_date: note!.modification_date || new Date().toISOString(),
            };
          }
        });
        
        // Add to database
        try {
          console.error(`Adding ${batchChunks.length} notes to database`);
          await notesTable.add(batchChunks);
          console.error("Successfully added batch to database");
          
          // Update progress
          state.processedNotes += batchChunks.length;
          await saveIndexingState(state);
        } catch (error) {
          const errorMsg = `Error adding batch to database: ${error.message}`;
          console.error(errorMsg);
          state.errors.push(errorMsg);
        }
      }
      
      // Small delay between batches to prevent overloading
      await new Promise(resolve => setTimeout(resolve, 500));
    }
    
    // Indexing complete
    state.inProgress = false;
    state.lastUpdateTime = Date.now();
    await saveIndexingState(state);
    console.error(`Indexing completed. Processed ${state.processedNotes} notes out of ${state.totalNotes}`);
  } catch (error) {
    // Handle any unexpected errors
    const errorMsg = `Unexpected error during indexing: ${error.message}`;
    console.error(errorMsg);
    state.errors.push(errorMsg);
    state.inProgress = false;
    state.lastUpdateTime = Date.now();
    await saveIndexingState(state);
  }
}

export const indexNotes = async (notesTable: any) => {
  const start = performance.now();
  let report = "";
  const allNotes = (await getNotes()) || [];
  const notesDetails = await Promise.all(
    allNotes.map((note) => {
      try {
        return getNoteDetailsByTitle(note);
      } catch (error) {
        report += `Error getting note details for ${note}: ${error.message}\n`;
        return {} as any;
      }
    })
  );

  const chunks = notesDetails
    .filter((n) => n.title)
    .map((node) => {
      try {
        return {
          ...node,
          content: turndown(node.content || ""), // this sometimes fails
        };
      } catch (error) {
        return node;
      }
    })
    .map((note, index) => ({
      id: index.toString(),
      title: note.title,
      content: note.content, // turndown(note.content || ""),
      creation_date: note.creation_date,
      modification_date: note.modification_date,
    }));

  await notesTable.add(chunks);

  return {
    chunks: chunks.length,
    report,
    allNotes: allNotes.length,
    time: performance.now() - start,
  };
};

export const createNotesTable = async (overrideName?: string) => {
  const start = performance.now();
  const notesTable = await db.createEmptyTable(
    overrideName || "notes",
    notesTableSchema,
    {
      mode: "create",
      existOk: true,
    }
  );

  const indices = await notesTable.listIndices();
  if (!indices.find((index) => index.name === "content_idx")) {
    await notesTable.createIndex("content", {
      config: lancedb.Index.fts(),
      replace: true,
    });
  }
  return { notesTable, time: performance.now() - start };
};

const createNote = async (title: string, content: string) => {
  // Escape special characters and convert newlines to \n
  const escapedTitle = title.replace(/[\\'"]/g, "\\$&");
  const escapedContent = content
    .replace(/[\\'"]/g, "\\$&")
    .replace(/\n/g, "\\n")
    .replace(/\r/g, "");

  await runJxa(`
    const app = Application('Notes');
    const note = app.make({new: 'note', withProperties: {
      name: "${escapedTitle}",
      body: "${escapedContent}"
    }});
    
    return true
  `);

  return true;
};

// Handle tool execution
server.setRequestHandler(CallToolRequestSchema, async (request, c) => {
  const { notesTable } = await createNotesTable();
  const { name, arguments: args } = request.params;

  try {
    if (name === "create-note") {
      const { title, content } = CreateNoteSchema.parse(args);
      await createNote(title, content);
      return createTextResponse(`Created note "${title}" successfully.`);
    } else if (name === "list-notes") {
      return createTextResponse(
        `There are ${await notesTable.countRows()} notes in your Apple Notes database.`
      );
    } else if (name == "get-note") {
      try {
        const { title } = GetNoteSchema.parse(args);
        const note = await getNoteDetailsByTitle(title);

        return createTextResponse(`${note}`);
      } catch (error) {
        return createTextResponse(error.message);
      }
    } else if (name === "index-notes") {
      // Start the background indexing process
      const status = await getIndexingStatus();
      
      if (status.inProgress) {
        return createTextResponse(
          `Indexing is already in progress. Progress: ${status.progress}% (${status.processedNotes}/${status.totalNotes} notes processed)`
        );
      }
      
      // Start the background indexing process
      backgroundIndexNotes(notesTable).catch(error => {
        console.error(`Background indexing error: ${error.message}`);
      });
      
      return createTextResponse(
        `Started indexing your Apple Notes in the background. This process will continue even if you close this chat. You can check the status by using the "index-notes" tool again.`
      );
    } else if (name === "search-notes") {
      const { query } = QueryNotesSchema.parse(args);
      const combinedResults = await searchAndCombineResults(notesTable, query);
      return createTextResponse(JSON.stringify(combinedResults));
    } else {
      throw new Error(`Unknown tool: ${name}`);
    }
  } catch (error) {
    if (error instanceof z.ZodError) {
      throw new Error(
        `Invalid arguments: ${error.errors
          .map((e) => `${e.path.join(".")}: ${e.message}`)
          .join(", ")}`
      );
    }
    throw error;
  }
});

// Start the server
const transport = new StdioServerTransport();
await server.connect(transport);
console.error("Local Machine MCP Server running on stdio");

const createTextResponse = (text: string) => ({
  content: [{ type: "text", text }],
});

/**
 * Search for notes by title or content using both vector and FTS search.
 * The results are combined using RRF
 */
export const searchAndCombineResults = async (
  notesTable: lancedb.Table,
  query: string,
  limit = 20
) => {
  const [vectorResults, ftsSearchResults] = await Promise.all([
    (async () => {
      const results = await notesTable
        .search(query, "vector")
        .limit(limit)
        .toArray();
      return results;
    })(),
    (async () => {
      const results = await notesTable
        .search(query, "fts", "content")
        .limit(limit)
        .toArray();
      return results;
    })(),
  ]);

  const k = 60;
  const scores = new Map<string, number>();

  const processResults = (results: any[], startRank: number) => {
    results.forEach((result, idx) => {
      const key = `${result.title}::${result.content}`;
      const score = 1 / (k + startRank + idx);
      scores.set(key, (scores.get(key) || 0) + score);
    });
  };

  processResults(vectorResults, 0);
  processResults(ftsSearchResults, 0);

  const results = Array.from(scores.entries())
    .sort(([, a], [, b]) => b - a)
    .slice(0, limit)
    .map(([key]) => {
      const [title, content] = key.split("::");
      return { title, content };
    });

  return results;
};

const CreateNoteSchema = z.object({
  title: z.string(),
  content: z.string(),
});
