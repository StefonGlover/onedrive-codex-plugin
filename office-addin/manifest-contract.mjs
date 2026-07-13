export const OFFICE_MANIFEST_HOSTS = Object.freeze({
  Document: "Word",
  Workbook: "Excel",
  Presentation: "PowerPoint"
});

export const OFFICE_ICON_SIZES = Object.freeze([16, 32, 64, 80]);
export const OFFICE_RIBBON_ICON_SIZES = Object.freeze([16, 32, 80]);

function countMatches(value, pattern) {
  return [...value.matchAll(pattern)].length;
}

function extractSingle(value, pattern, description, problems) {
  const matches = [...value.matchAll(pattern)];
  if (matches.length !== 1) {
    problems.push(`Office manifest must contain exactly one ${description}; found ${matches.length}.`);
    return null;
  }
  return matches[0];
}

function requireText(value, text, description, problems) {
  if (!value.includes(text)) problems.push(`Office manifest is missing ${description}.`);
}

function checkRibbonIcons(value, location, problems) {
  const images = [...value.matchAll(/<bt:Image\s+size="(\d+)"\s+resid="([^"]+)"\s*\/>/g)]
    .map((match) => ({ size: Number(match[1]), resid: match[2] }));
  const expected = OFFICE_RIBBON_ICON_SIZES.map((size) => ({ size, resid: `Icon.${size}` }));
  if (JSON.stringify(images) !== JSON.stringify(expected)) {
    problems.push(`${location} must reference exactly the 16, 32, and 80 pixel ribbon icons in ascending order.`);
  }
}

function checkHost(hostType, hostBody, problems) {
  const product = OFFICE_MANIFEST_HOSTS[hostType];
  const formFactor = extractSingle(
    hostBody,
    /<DesktopFormFactor>([\s\S]*?)<\/DesktopFormFactor>/g,
    `${product} DesktopFormFactor`,
    problems
  )?.[1];
  if (!formFactor) return;

  const getStarted = extractSingle(formFactor, /<GetStarted>([\s\S]*?)<\/GetStarted>/g, `${product} GetStarted`, problems)?.[1] || "";
  requireText(getStarted, '<Title resid="GetStarted.Title"/>', `${product} GetStarted title`, problems);
  requireText(getStarted, '<Description resid="GetStarted.Description"/>', `${product} GetStarted description`, problems);
  requireText(getStarted, '<LearnMoreUrl resid="LearnMore.Url"/>', `${product} GetStarted learn-more URL`, problems);

  if (countMatches(formFactor, /<FunctionFile\s+resid="Function\.Url"\s*\/>/g) !== 1) {
    problems.push(`${product} must contain exactly one FunctionFile referencing Function.Url.`);
  }
  const extension = extractSingle(
    formFactor,
    /<ExtensionPoint\s+xsi:type="PrimaryCommandSurface">([\s\S]*?)<\/ExtensionPoint>/g,
    `${product} PrimaryCommandSurface`,
    problems
  )?.[1] || "";
  const officeTab = extractSingle(extension, /<OfficeTab\s+id="TabHome">([\s\S]*?)<\/OfficeTab>/g, `${product} Home ribbon tab`, problems)?.[1] || "";
  const groupId = `CodexOneDrive.${product}.Group`;
  const group = extractSingle(
    officeTab,
    new RegExp(`<Group\\s+id="${groupId.replaceAll(".", "\\.")}">([\\s\\S]*?)<\\/Group>`, "g"),
    `${product} ribbon group`,
    problems
  )?.[1] || "";
  requireText(group, '<Label resid="Group.Label"/>', `${product} ribbon group label`, problems);

  const icons = [...group.matchAll(/<Icon>([\s\S]*?)<\/Icon>/g)];
  if (icons.length !== 2) {
    problems.push(`${product} must contain exactly one group icon and one button icon; found ${icons.length}.`);
  } else {
    checkRibbonIcons(icons[0][1], `${product} ribbon group icon`, problems);
    checkRibbonIcons(icons[1][1], `${product} ribbon button icon`, problems);
  }

  const buttonId = `CodexOneDrive.${product}.ShowTaskpane`;
  const button = extractSingle(
    group,
    new RegExp(`<Control\\s+xsi:type="Button"\\s+id="${buttonId.replaceAll(".", "\\.")}">([\\s\\S]*?)<\\/Control>`, "g"),
    `${product} ShowTaskpane button`,
    problems
  )?.[1] || "";
  requireText(button, '<Label resid="Button.Label"/>', `${product} button label`, problems);
  requireText(button, '<Title resid="Button.Label"/>', `${product} button supertip title`, problems);
  requireText(button, '<Description resid="Button.Tooltip"/>', `${product} button supertip description`, problems);

  const action = extractSingle(
    button,
    /<Action\s+xsi:type="ShowTaskpane">([\s\S]*?)<\/Action>/g,
    `${product} ShowTaskpane action`,
    problems
  )?.[1] || "";
  requireText(action, "<TaskpaneId>CodexOneDriveTaskpane</TaskpaneId>", `${product} task-pane ID`, problems);
  requireText(action, '<SourceLocation resid="Taskpane.Url"/>', `${product} task-pane source location`, problems);

  const getStartedIndex = formFactor.indexOf("<GetStarted>");
  const functionIndex = formFactor.indexOf("<FunctionFile");
  const extensionIndex = formFactor.indexOf("<ExtensionPoint");
  if (!(getStartedIndex >= 0 && getStartedIndex < functionIndex && functionIndex < extensionIndex)) {
    problems.push(`${product} DesktopFormFactor children must be ordered GetStarted, FunctionFile, ExtensionPoint.`);
  }
}

function resourceDefinitions(resources, container, tag) {
  const body = resources.match(new RegExp(`<bt:${container}>([\\s\\S]*?)<\\/bt:${container}>`))?.[1] || "";
  return [...body.matchAll(new RegExp(`<bt:${tag}\\s+id="([^"]+)"\\s+DefaultValue="([^"]+)"\\s*\\/>`, "g"))]
    .map((match) => ({ id: match[1], value: match[2] }));
}

export function officeManifestProblems(manifest) {
  const problems = [];
  requireText(manifest, "<Version>1.1.1.0</Version>", "Office companion version 1.1.1.0", problems);
  requireText(manifest, 'xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"', "task-pane override namespace", problems);
  requireText(manifest, 'IconUrl DefaultValue="https://127.0.0.1:3443/office-addin/icon-32.png"', "base 32 pixel icon URL", problems);
  requireText(manifest, 'HighResolutionIconUrl DefaultValue="https://127.0.0.1:3443/office-addin/icon-64.png"', "base 64 pixel icon URL", problems);

  const overrideMatch = extractSingle(
    manifest,
    /<VersionOverrides\s+xmlns="http:\/\/schemas\.microsoft\.com\/office\/taskpaneappversionoverrides"\s+xsi:type="VersionOverridesV1_0">([\s\S]*?)<\/VersionOverrides>/g,
    "VersionOverridesV1_0 block",
    problems
  );
  if (!overrideMatch) return problems;
  const override = overrideMatch[1];
  if (countMatches(override, /<bt:Set\s+Name="AddinCommands"\s*\/>/g) !== 1
    || !override.includes('<bt:Sets DefaultMinVersion="1.1">')) {
    problems.push("Office VersionOverrides must require AddinCommands 1.1 exactly once.");
  }
  if (override.includes("ExecuteFunction")) {
    problems.push("Office ribbon command must not use ExecuteFunction.");
  }

  const hostMatches = [...override.matchAll(/<Host\s+xsi:type="([^"]+)">([\s\S]*?)<\/Host>/g)];
  const actualHosts = hostMatches.map((match) => match[1]);
  const expectedHosts = Object.keys(OFFICE_MANIFEST_HOSTS);
  if (JSON.stringify(actualHosts) !== JSON.stringify(expectedHosts)) {
    problems.push(`Office VersionOverrides hosts must be exactly ${expectedHosts.join(", ")}; found ${actualHosts.join(", ") || "none"}.`);
  }
  for (const match of hostMatches) {
    if (Object.hasOwn(OFFICE_MANIFEST_HOSTS, match[1])) checkHost(match[1], match[2], problems);
  }

  if (countMatches(override, /<Action\s+xsi:type="ShowTaskpane">/g) !== expectedHosts.length) {
    problems.push(`Office manifest must contain exactly ${expectedHosts.length} ShowTaskpane actions.`);
  }
  const resources = extractSingle(override, /<Resources>([\s\S]*?)<\/Resources>/g, "Resources block", problems)?.[1] || "";
  const images = resourceDefinitions(resources, "Images", "Image");
  const urls = resourceDefinitions(resources, "Urls", "Url");
  const shortStrings = resourceDefinitions(resources, "ShortStrings", "String");
  const longStrings = resourceDefinitions(resources, "LongStrings", "String");
  const expectedImages = OFFICE_RIBBON_ICON_SIZES.map((size) => ({
    id: `Icon.${size}`,
    value: `https://127.0.0.1:3443/office-addin/icon-${size}.png`
  }));
  const expectedUrls = [
    { id: "Function.Url", value: "https://127.0.0.1:3443/office-addin/taskpane.html" },
    { id: "Taskpane.Url", value: "https://127.0.0.1:3443/office-addin/taskpane.html" },
    { id: "LearnMore.Url", value: "https://github.com/StefonGlover/onedrive-codex-plugin" }
  ];
  if (JSON.stringify(images) !== JSON.stringify(expectedImages)) {
    problems.push("Office manifest image resources must be exactly the local 16, 32, and 80 pixel PNGs.");
  }
  if (JSON.stringify(urls) !== JSON.stringify(expectedUrls)) {
    problems.push("Office manifest URL resources must contain the exact function, task-pane, and learn-more URLs.");
  }
  if (JSON.stringify(shortStrings.map(({ id }) => id)) !== JSON.stringify(["GetStarted.Title", "Group.Label", "Button.Label"])) {
    problems.push("Office manifest ShortStrings must contain the exact GetStarted, group, and button resource IDs.");
  }
  if (JSON.stringify(longStrings.map(({ id }) => id)) !== JSON.stringify(["GetStarted.Description", "Button.Tooltip"])) {
    problems.push("Office manifest LongStrings must contain the exact GetStarted description and button tooltip resource IDs.");
  }

  const definitions = [...images, ...urls, ...shortStrings, ...longStrings];
  const definitionIds = definitions.map(({ id }) => id);
  if (new Set(definitionIds).size !== definitionIds.length) {
    problems.push("Office manifest resource IDs must be unique.");
  }
  for (const { id, value } of definitions) {
    if (id.length > 32) problems.push(`Office manifest resource ID exceeds 32 characters: ${id}.`);
    if (urls.some((resource) => resource.id === id) && !value.startsWith("https://")) {
      problems.push(`Office manifest URL resource must use HTTPS: ${id}.`);
    }
  }
  for (const { value } of shortStrings) {
    if (value.length > 125) problems.push("Office manifest ShortString exceeds 125 characters.");
  }
  for (const { value } of longStrings) {
    if (value.length > 250) problems.push("Office manifest LongString exceeds 250 characters.");
  }

  const referencedIds = [...override.matchAll(/\bresid="([^"]+)"/g)].map((match) => match[1]);
  for (const resid of new Set(referencedIds)) {
    if (!definitionIds.includes(resid)) problems.push(`Office manifest references undefined resource ID: ${resid}.`);
  }
  for (const id of definitionIds) {
    const referenceCount = referencedIds.filter((resid) => resid === id).length;
    if (referenceCount === 0) problems.push(`Office manifest defines unused resource ID: ${id}.`);
  }
  return problems;
}
