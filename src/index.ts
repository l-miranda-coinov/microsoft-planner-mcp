import { FastMCP } from "fastmcp";
import { z } from "zod";
import { execFileSync } from "child_process";

const mcp = new FastMCP({
  name: "microsoft-planner-mcp",
  version: "1.0.0",
});

const GRAPH = "https://graph.microsoft.com/v1.0";

// IDs do Planner/Graph: apenas letras, números, "-" e "_" (evita injeção em URL/argumentos)
const idSchema = (description: string) =>
  z
    .string()
    .regex(/^[A-Za-z0-9_-]{1,128}$/, "Invalid ID format")
    .describe(description);

const categorySchema = z
  .string()
  .regex(/^category([1-9]|1[0-9]|2[0-5])$/, "Must be category1 to category25");

// Executa `az` SEM shell: cada argumento é passado de forma literal (sem interpretação de shell)
function runAz(args: string[]): string {
  try {
    return execFileSync("az", args, {
      encoding: "utf-8",
      maxBuffer: 10 * 1024 * 1024,
      stdio: ["ignore", "pipe", "pipe"],
    });
  } catch (error: any) {
    // Não ecoa o comando completo; devolve só a primeira linha do stderr
    const stderr = String(error?.stderr ?? "").trim().split("\n")[0];
    throw new Error(`az rest failed${stderr ? `: ${stderr}` : ""}`);
  }
}

// Helper to execute az rest commands
function azRest(
  method: string,
  url: string,
  body?: object,
  headers: string[] = []
): string {
  const args = ["rest", "--method", method, "--url", url];
  const allHeaders = [...headers];
  if (body) allHeaders.unshift("Content-Type=application/json");
  if (allHeaders.length) args.push("--headers", ...allHeaders);
  if (body) args.push("--body", JSON.stringify(body));
  return runAz(args);
}

// Helper to get ETag for update/delete operations
function getETag(taskId: string, isDetails: boolean = false): string {
  const url = isDetails
    ? `${GRAPH}/planner/tasks/${taskId}/details`
    : `${GRAPH}/planner/tasks/${taskId}`;
  const result = JSON.parse(azRest("GET", url));
  const etag = result["@odata.etag"];
  if (typeof etag !== "string") throw new Error("ETag not found in response");
  return etag;
}

// Tool: List tasks for a plan
mcp.addTool({
  name: "list-tasks",
  description: "List all tasks in a Planner plan",
  parameters: z.object({
    planId: idSchema("The Planner plan ID"),
  }),
  execute: async ({ planId }) => {
    const result = JSON.parse(azRest("GET", `${GRAPH}/planner/plans/${planId}/tasks`));
    return JSON.stringify(result.value, null, 2);
  },
});

// Tool: Get single task
mcp.addTool({
  name: "get-task",
  description: "Get details of a specific Planner task",
  parameters: z.object({
    taskId: idSchema("The task ID"),
  }),
  execute: async ({ taskId }) => azRest("GET", `${GRAPH}/planner/tasks/${taskId}`),
});

// Tool: Get task details (description, checklist, references)
mcp.addTool({
  name: "get-task-details",
  description: "Get extended task details including description and checklist",
  parameters: z.object({
    taskId: idSchema("The task ID"),
  }),
  execute: async ({ taskId }) => azRest("GET", `${GRAPH}/planner/tasks/${taskId}/details`),
});

// Tool: Create task
mcp.addTool({
  name: "create-task",
  description: "Create a new task in a Planner plan",
  parameters: z.object({
    planId: idSchema("The plan ID"),
    bucketId: idSchema("The bucket ID"),
    title: z.string().min(1).max(256).describe("Task title"),
  }),
  execute: async ({ planId, bucketId, title }) =>
    azRest("POST", `${GRAPH}/planner/tasks`, { planId, bucketId, title }),
});

// Tool: Update task (title, percentComplete, assignments, categories)
mcp.addTool({
  name: "update-task",
  description: "Update task properties (title, progress, assignments, categories). Auto-fetches ETag.",
  parameters: z.object({
    taskId: idSchema("The task ID"),
    title: z.string().min(1).max(256).optional().describe("New title"),
    percentComplete: z.number().min(0).max(100).optional().describe("Progress 0-100"),
    assignUserId: idSchema("User ID to assign").optional(),
    category: categorySchema.optional().describe("Category to apply (category1-category25)"),
  }),
  execute: async ({ taskId, title, percentComplete, assignUserId, category }) => {
    const etag = getETag(taskId);
    const body: Record<string, any> = {};
    if (title !== undefined) body.title = title;
    if (percentComplete !== undefined) body.percentComplete = percentComplete;
    if (assignUserId) {
      body.assignments = {
        [assignUserId]: {
          "@odata.type": "#microsoft.graph.plannerAssignment",
          orderHint: " !",
        },
      };
    }
    if (category) {
      body.appliedCategories = { [category]: true };
    }

    const result = azRest("PATCH", `${GRAPH}/planner/tasks/${taskId}`, body, [
      `If-Match=${etag}`,
    ]);
    return result || "Task updated successfully";
  },
});

// Tool: Update task details (description with GitHub links)
mcp.addTool({
  name: "update-task-details",
  description: "Update task description (use for GitHub links). Auto-fetches ETag.",
  parameters: z.object({
    taskId: idSchema("The task ID"),
    description: z
      .string()
      .max(32768)
      .describe("Task description (supports markdown, include GitHub URLs)"),
  }),
  execute: async ({ taskId, description }) => {
    const etag = getETag(taskId, true);
    const result = azRest(
      "PATCH",
      `${GRAPH}/planner/tasks/${taskId}/details`,
      { description },
      [`If-Match=${etag}`]
    );
    return result || "Task details updated successfully";
  },
});

// Tool: Delete task
mcp.addTool({
  name: "delete-task",
  description: "Delete a Planner task. Auto-fetches ETag.",
  parameters: z.object({
    taskId: idSchema("The task ID to delete"),
  }),
  execute: async ({ taskId }) => {
    const etag = getETag(taskId);
    azRest("DELETE", `${GRAPH}/planner/tasks/${taskId}`, undefined, [`If-Match=${etag}`]);
    return "Task deleted successfully";
  },
});

// Tool: List buckets for a plan
mcp.addTool({
  name: "list-buckets",
  description: "List all buckets in a Planner plan",
  parameters: z.object({
    planId: idSchema("The Planner plan ID"),
  }),
  execute: async ({ planId }) => {
    const result = JSON.parse(azRest("GET", `${GRAPH}/planner/plans/${planId}/buckets`));
    return JSON.stringify(result.value, null, 2);
  },
});

// Tool: List plans for current user
mcp.addTool({
  name: "list-plans",
  description: "List all Planner plans accessible to the current user",
  parameters: z.object({}),
  execute: async () => {
    const result = JSON.parse(azRest("GET", `${GRAPH}/me/planner/plans`));
    return JSON.stringify(result.value, null, 2);
  },
});

mcp.start({ transportType: "stdio" });
