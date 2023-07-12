# Variables Evaluator Extension

Allows to use variables in your messages and replace them using a scripted rule.

## How to Use

Compile and put resulting dlls to the Extension folder of the SV Designer/Server. 

E.g. `c:\Program Files\Micro Focus\Service Virtualization Designer\Designer\Extensions\`

## How to Compile

Create `SV_BIN_FOLDER` environment variable and point it to the SV bin folder. 

E.g. `c:\Program Files\Micro Focus\Service Virtualization Designer\Designer\bin\`

## Usage Example

```
// Populate dictionary with variables
Dictionary<string, string> variables = new Dictionary<string, string>();
variables.Add("title", "This is a title");
variables.Add("description", "This is a description");
variables.Add("firstTag", "tag1");
variables.Add("secondTag", "tag2");

// replaces variables in the sv.Response.JSONResponse.Book subtree. Variables must be using %% at the start and end. E.g. %%title%%
sv.Response.JSONResponse.Book.EvaluateVariables(variables, sv.MessageLogger);

// if you are using different variable start/end sequence than %%, you must explicitly specify it. E.g. if you are using brackets like {title}
sv.Response.JSONResponse.Article.EvaluateVariables(variables, "{", "}");
```

