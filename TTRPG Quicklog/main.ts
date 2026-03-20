import {
  App,
  Editor,
  EditorPosition,
  EditorSuggest,
  EditorSuggestContext,
  EditorSuggestTriggerInfo,
  Modal,
  MarkdownView,
  Notice,
  Plugin,
  PluginSettingTab,
  Setting,
  TFile,
  TFolder,
  normalizePath,
} from "obsidian";
import { EditorState, RangeSetBuilder, type Extension } from "@codemirror/state";
import { Decoration, type DecorationSet, EditorView, ViewPlugin, type ViewUpdate } from "@codemirror/view";
import { autocompletion, type Completion, type CompletionContext, type CompletionResult } from "@codemirror/autocomplete";

interface MentionTypeConfig {
  prefix: string;
  entityType: string;
  color?: string;
}

interface CampaignConfig {
  id: string;
  name: string;
  root: string;
  entityFolders: Record<string, string>;
}

interface CampaignNamespaceLinkerSettings {
  campaigns: CampaignConfig[];
  sharedRoots: string[];
  includeSharedInSuggestions: boolean;
  enableProseSuggestions: boolean;
  proseSuggestionMinChars: number;
  mentionTypes: MentionTypeConfig[];
  mentionInputModes: MentionInputModes;
  capitalizeDefaultCategoryFolders: boolean;
  useRainbowCategoryColors: boolean;
  prefixOverlayHotkey: string;
}

interface MentionInputModes {
  singleWord: boolean;
  quoted: boolean;
  open: boolean;
}

interface CategoryRenameDecision {
  applyTypeUpdate: boolean;
  renameFolders: boolean;
}

interface VariantMatch {
  variant: string;
  count: number;
  files: Set<string>;
}

interface VariantBackfillDecision {
  applyLinks: boolean;
  addAliases: boolean;
  selectedVariants: string[];
}

interface CategorizeDocumentDecision {
  entityType: string;
  aliases: string[];
}

type VariantMatchExpander = (customSpellings: string[]) => Promise<VariantMatch[]>;

interface ProseMentionSuggestion {
  record: EntityRecord;
  displayText: string;
  normalizedDisplay: string;
  isAlias: boolean;
}

interface EntityRecord {
  file: TFile;
  name: string;
  normalizedName: string;
  aliases: string[];
  normalizedAliases: string[];
  entityType: string;
  scope: "campaign" | "shared";
  campaignId: string | null;
}

interface CampaignScope {
  campaign: CampaignConfig | null;
  inShared: boolean;
}

interface ParsedMentionQuery {
  prefix: string;
  entityType: string;
  rawQuery: string;
}

type MentionSuggestion =
  | {
      kind: "existing";
      record: EntityRecord;
    }
  | {
      kind: "create";
      name: string;
      entityType: string;
      campaign: CampaignConfig;
    };

const RAINBOW_COLORS = [
  "#e53935",
  "#ef5350",
  "#fb8c00",
  "#ffb300",
  "#fdd835",
  "#43a047",
  "#00acc1",
  "#1e88e5",
  "#5e35b1",
  "#8e24aa",
];

const DEFAULT_SETTINGS: CampaignNamespaceLinkerSettings = {
  campaigns: [],
  sharedRoots: [],
  includeSharedInSuggestions: true,
  enableProseSuggestions: true,
  proseSuggestionMinChars: 2,
  mentionTypes: [
    { prefix: "==", entityType: "pcs" },
    { prefix: "@@", entityType: "npcs" },
    { prefix: "@@@", entityType: "allies" },
    { prefix: "@@@@", entityType: "adversaries" },
    { prefix: "##", entityType: "locations" },
    { prefix: "&&", entityType: "factions" },
    { prefix: "$$", entityType: "loot" },
    { prefix: "!!", entityType: "magic items" },
    { prefix: "??", entityType: "quests" },
    { prefix: "%%", entityType: "events" },
  ],
  mentionInputModes: {
    singleWord: true,
    quoted: true,
    open: true,
  },
  capitalizeDefaultCategoryFolders: true,
  useRainbowCategoryColors: true,
  prefixOverlayHotkey: "mod+shift+o",
};

export default class CampaignNamespaceLinkerPlugin extends Plugin {
  settings: CampaignNamespaceLinkerSettings = DEFAULT_SETTINGS;
  private entities: EntityRecord[] = [];
  private suggester: CampaignMentionSuggester | null = null;
  private prefixOverlayEl: HTMLElement | null = null;
  private completionColorStyleEl: HTMLStyleElement | null = null;

  async onload(): Promise<void> {
    await this.loadSettings();
    await this.rebuildEntityIndex();
    this.refreshCompletionColorStyles();

    this.suggester = new CampaignMentionSuggester(this);
    this.registerEditorSuggest(this.suggester);
    this.registerEditorExtension(this.createProseAutocompleteExtension());
    this.registerEditorExtension(this.createEditorLinkColorExtension());

    this.registerEvent(this.app.vault.on("create", () => this.queueReindex()));
    this.registerEvent(this.app.vault.on("modify", () => this.queueReindex()));
    this.registerEvent(this.app.vault.on("rename", () => this.queueReindex()));
    this.registerEvent(this.app.vault.on("delete", () => this.queueReindex()));
    this.registerEvent(this.app.metadataCache.on("resolved", () => this.queueReindex()));
    this.registerEvent(this.app.workspace.on("active-leaf-change", () => this.queueLinkColorRefresh()));
    this.registerEvent(this.app.workspace.on("layout-change", () => this.queueLinkColorRefresh()));
    this.registerDomEvent(document, "keydown", (evt: KeyboardEvent) => this.handlePrefixOverlayHotkey(evt));

    this.addSettingTab(new CampaignNamespaceLinkerSettingTab(this.app, this));
    this.registerMarkdownPostProcessor((el, ctx) => {
      this.applyRenderedLinkColors(el, ctx.sourcePath);
    });

    this.addCommand({
      id: "register-parent-folder-as-campaign-root",
      name: "Register parent folder as campaign root",
      callback: async () => {
        const file = this.getActiveFile();
        if (!file) {
          new Notice("No active file.");
          return;
        }

        const parts = file.path.split("/");
        if (parts.length < 2) {
          new Notice("Active file is at vault root. Move it into a campaign folder first.");
          return;
        }
        parts.pop();
        const root = normalizePath(parts.join("/"));
        const name = parts[parts.length - 1] ?? "Campaign";
        const id = slugify(name);

        if (this.settings.campaigns.some((c) => c.root === root)) {
          new Notice("That folder is already registered as a campaign root.");
          return;
        }

        const campaign: CampaignConfig = {
          id,
          name,
          root,
          entityFolders: {},
        };
        this.settings.campaigns.push(campaign);
        await this.saveSettings();
        new Notice(`Registered campaign root: ${root}`);
        await this.offerCreateCategoryFoldersForCampaign(campaign);
      },
    });

    this.addCommand({
      id: "convert-selected-mentions-to-links",
      name: "Convert selected mentions to links",
      editorCallback: async (editor, view) => {
        const file = view.file;
        if (!file) {
          new Notice("No active file.");
          return;
        }

        const selection = editor.getSelection();
        if (!selection) {
          new Notice("Select text that contains mentions.");
          return;
        }

        const scope = this.resolveScopeForFile(file);
        if (!scope.campaign && !scope.inShared) {
          new Notice("This note is outside configured campaign/shared roots.");
          return;
        }

        const converted = this.convertMentionsInText(selection, file, scope);
        editor.replaceSelection(converted);
      },
    });

    this.addCommand({
      id: "update-active-entity-from-campaign-prose",
      name: "Update active entity aliases from campaign prose",
      callback: async () => {
        const file = this.getActiveFile();
        if (!file) {
          new Notice("No active file.");
          return;
        }

        const scope = this.resolveScopeForFile(file);
        if (!scope.campaign) {
          new Notice("Active note is outside configured campaign roots.");
          return;
        }

        const knownEntity = this.entities.some(
          (record) => record.file.path === file.path && record.scope === "campaign" && record.campaignId === scope.campaign?.id,
        );
        if (!knownEntity) {
          new Notice("Active note is not indexed as a campaign entity.");
          return;
        }

        await this.updateEntityFromCampaignProse(file, scope.campaign);
      },
    });

    this.addCommand({
      id: "categorize-active-document-as-entity",
      name: "Categorize active document as entity",
      callback: async () => {
        const file = this.getActiveFile();
        if (!file) {
          new Notice("No active file.");
          return;
        }

        const scope = this.resolveScopeForFile(file);
        if (!scope.campaign) {
          new Notice("Active note is outside configured campaign roots.");
          return;
        }

        const knownTypes = this.getKnownEntityTypes();
        if (knownTypes.length === 0) {
          new Notice("No category types are configured.");
          return;
        }

        const cache = this.app.metadataCache.getFileCache(file);
        const fm = cache?.frontmatter;
        const hasType = typeof fm?.type === "string" && fm.type.trim().length > 0;
        const hasCampaign = typeof fm?.campaign === "string" && fm.campaign.trim().length > 0;
        if (hasType || hasCampaign) {
          new Notice("Active note is already categorized.");
          return;
        }

        const decision = await CategorizeDocumentModal.prompt(this.app, file.basename, knownTypes);
        if (!decision) return;

        const canonicalName = file.basename;
        const existingAliases = frontmatterAliasesToArray(fm?.aliases);
        const aliases = Array.from(
          new Set(
            [...existingAliases, ...decision.aliases]
              .map((alias) => normalizeWhitespace(alias))
              .filter((alias) => alias.length > 0 && normalizeSearchText(alias) !== normalizeSearchText(canonicalName)),
          ),
        );

        await this.app.fileManager.processFrontMatter(file, (frontmatter) => {
          frontmatter.type = normalizeEntityType(decision.entityType);
          frontmatter.campaign = scope.campaign?.id;
          frontmatter.aliases = aliases;
        });

        await this.rebuildEntityIndex();
        this.queueLinkColorRefresh();
        await this.maybeOfferVariantBackfill(file, canonicalName, scope.campaign, aliases);
        new Notice(`Categorized ${canonicalName} as ${normalizeEntityType(decision.entityType)}.`);
      },
    });

    this.queueLinkColorRefresh();
  }

  async onunload(): Promise<void> {
    this.hidePrefixOverlay();
    this.completionColorStyleEl?.remove();
    this.completionColorStyleEl = null;
    this.entities = [];
    this.suggester = null;
  }

  async loadSettings(): Promise<void> {
    const loaded = await this.loadData();
    this.settings = Object.assign({}, DEFAULT_SETTINGS, loaded);
    this.settings.mentionInputModes = Object.assign(
      {},
      DEFAULT_SETTINGS.mentionInputModes,
      loaded?.mentionInputModes ?? {},
    );
    this.settings.capitalizeDefaultCategoryFolders =
      loaded?.capitalizeDefaultCategoryFolders ?? DEFAULT_SETTINGS.capitalizeDefaultCategoryFolders;
    this.settings.useRainbowCategoryColors =
      loaded?.useRainbowCategoryColors ?? DEFAULT_SETTINGS.useRainbowCategoryColors;
    this.settings.enableProseSuggestions =
      loaded?.enableProseSuggestions ?? DEFAULT_SETTINGS.enableProseSuggestions;
    this.settings.proseSuggestionMinChars = clampNumber(
      loaded?.proseSuggestionMinChars,
      1,
      10,
      DEFAULT_SETTINGS.proseSuggestionMinChars,
    );
    this.settings.prefixOverlayHotkey = normalizeHotkeySetting(
      loaded?.prefixOverlayHotkey ?? DEFAULT_SETTINGS.prefixOverlayHotkey,
    );

    this.settings.sharedRoots = this.settings.sharedRoots.map((p) => normalizePath(p));
    this.settings.campaigns = this.settings.campaigns.map((c) => ({
      id: c.id || slugify(c.name),
      name: c.name,
      root: normalizePath(c.root),
      entityFolders: c.entityFolders || {},
    }));
    this.settings.mentionTypes = this.settings.mentionTypes
      .filter((m) => m.prefix.trim().length > 0 && m.entityType.trim().length > 0)
      .map((m) => ({
        prefix: m.prefix,
        entityType: normalizeEntityType(m.entityType),
        color: normalizeHexColor(m.color),
      }));

    if (this.settings.mentionTypes.length === 0) {
      this.settings.mentionTypes = DEFAULT_SETTINGS.mentionTypes;
    }
  }

  async saveSettings(): Promise<void> {
    await this.saveData(this.settings);
    await this.rebuildEntityIndex();
    this.refreshCompletionColorStyles();
    this.queueLinkColorRefresh();
    this.forceEditorColorRefresh();
    if (this.prefixOverlayEl) {
      this.renderPrefixOverlay();
    }
  }

  getActiveFile(): TFile | null {
    return this.app.workspace.getActiveFile();
  }

  queueReindex(): void {
    window.setTimeout(() => {
      void this.rebuildEntityIndex().then(() => {
        this.queueLinkColorRefresh();
      });
    }, 0);
  }

  queueLinkColorRefresh(): void {
    window.setTimeout(() => {
      const view = this.app.workspace.getActiveViewOfType(MarkdownView);
      const file = view?.file;
      const renderedContainer = view?.contentEl?.querySelector<HTMLElement>(".markdown-preview-view");
      if (renderedContainer && file) {
        this.applyRenderedLinkColors(renderedContainer, file.path);
      }
    }, 0);
  }

  forceEditorColorRefresh(): void {
    const view = this.app.workspace.getActiveViewOfType(MarkdownView);
    const cm = (view?.editor as Editor & { cm?: EditorView })?.cm;
    if (!cm) return;
    cm.dispatch({ selection: cm.state.selection });
  }

  getMentionTypeForPrefix(prefix: string): MentionTypeConfig | null {
    return this.settings.mentionTypes.find((m) => m.prefix === prefix) ?? null;
  }

  getEffectiveColorForType(entityType: string): string | null {
    const normalizedType = this.canonicalizeEntityType(entityType);
    if (!normalizedType) return null;
    const direct = this.settings.mentionTypes.find((m) => normalizeEntityType(m.entityType) === normalizedType);
    if (direct?.color) return direct.color;
    if (!this.settings.useRainbowCategoryColors) return null;

    const orderedTypes = Array.from(
      new Set(this.settings.mentionTypes.map((m) => normalizeEntityType(m.entityType)).filter(Boolean)),
    );
    const idx = orderedTypes.indexOf(normalizedType);
    if (idx < 0) return null;
    return RAINBOW_COLORS[idx % RAINBOW_COLORS.length] ?? null;
  }

  resolveEntityTypeForLink(linkText: string, sourcePath: string): string | null {
    const normalizedLink = normalizeLinkPathForResolve(linkText);
    if (!normalizedLink) return null;
    const destination = this.app.metadataCache.getFirstLinkpathDest(normalizedLink, sourcePath);
    if (!(destination instanceof TFile)) return null;

    const indexed = this.entities.find((e) => e.file.path === destination.path)?.entityType;
    if (indexed) {
      return this.canonicalizeEntityType(indexed);
    }

    const typeRaw = this.app.metadataCache.getFileCache(destination)?.frontmatter?.type;
    if (typeof typeRaw === "string") {
      return this.canonicalizeEntityType(typeRaw);
    }

    const campaign = this.resolveCampaignForPath(destination.path);
    if (campaign) {
      const fromFolder = this.inferEntityTypeFromCampaignFolder(destination.path, campaign);
      if (fromFolder) {
        return this.canonicalizeEntityType(fromFolder);
      }
    }
    const scope = this.resolveScopeForFile(destination);
    if (scope.inShared) {
      const fromSharedFolder = this.inferEntityTypeFromSharedFolder(destination.path);
      if (fromSharedFolder) {
        return this.canonicalizeEntityType(fromSharedFolder);
      }
    }
    return null;
  }

  canonicalizeEntityType(entityType: string): string | null {
    const normalized = normalizeEntityType(entityType);
    if (!normalized) return null;

    const known = new Set(this.settings.mentionTypes.map((m) => normalizeEntityType(m.entityType)));
    if (known.has(normalized)) return normalized;

    if (normalized.endsWith("s")) {
      const singular = normalized.slice(0, -1);
      if (known.has(singular)) return singular;
    } else {
      const plural = `${normalized}s`;
      if (known.has(plural)) return plural;
    }

    return null;
  }

  applyRenderedLinkColors(container: HTMLElement, sourcePath: string): void {
    const links = container.querySelectorAll<HTMLAnchorElement>("a.internal-link");
    links.forEach((linkEl) => {
      const href = linkEl.getAttribute("data-href") ?? linkEl.getAttribute("href") ?? "";
      if (!href) return;
      const type = this.resolveEntityTypeForLink(href, sourcePath);
      const color = type ? this.getEffectiveColorForType(type) : null;
      if (color && type) {
        this.applyCategoryColorToElement(linkEl, color, type);
      } else {
        this.clearCategoryColorFromElement(linkEl);
      }
    });
  }

  createEditorLinkColorExtension(): Extension {
    const plugin = this;
    const wikilinkPattern = /\[\[([^\]\n]+)\]\]/g;

    return ViewPlugin.fromClass(
      class {
        decorations: DecorationSet;

        constructor(view: EditorView) {
          this.decorations = this.buildDecorations(view);
        }

        update(update: ViewUpdate): void {
          if (update.docChanged || update.viewportChanged || update.selectionSet) {
            this.decorations = this.buildDecorations(update.view);
          }
        }

        buildDecorations(view: EditorView): DecorationSet {
          const builder = new RangeSetBuilder<Decoration>();
          const sourcePath = plugin.app.workspace.getActiveFile()?.path ?? "";

          for (const { from, to } of view.visibleRanges) {
            const slice = view.state.doc.sliceString(from, to);
            wikilinkPattern.lastIndex = 0;
            let match: RegExpExecArray | null;

            while ((match = wikilinkPattern.exec(slice)) !== null) {
              const inner = match[1];
              if (!inner) continue;
              const linkPath = normalizeLinkPathForResolve(inner);
              if (!linkPath) continue;

              const type = plugin.resolveEntityTypeForLink(linkPath, sourcePath);
              const color = type ? plugin.getEffectiveColorForType(type) : null;
              if (!type || !color) continue;

              const start = from + match.index;
              const end = start + match[0].length;
              const mark = Decoration.mark({
                class: "ttrpg-quicklog-colored-link",
                attributes: {
                  style: `--ttrpg-link-color: ${color}; color: ${color} !important;`,
                  "data-ttrpg-type": type,
                },
              });
              builder.add(start, end, mark);
            }
          }

          return builder.finish();
        }
      },
      {
        decorations: (value) => value.decorations,
      },
    );
  }

  createProseAutocompleteExtension(): Extension {
    const plugin = this;
    const source = (context: CompletionContext): CompletionResult | null => {
      if (!plugin.settings.enableProseSuggestions) return null;
      const activeFile = plugin.getActiveFile();
      if (!activeFile) return null;
      const scope = plugin.resolveScopeForFile(activeFile);
      if (!scope.campaign) return null;

      const line = context.state.doc.lineAt(context.pos);
      const before = line.text.slice(0, context.pos - line.from);
      if (/\[\[[^\]]*$/.test(before) || /`[^`]*$/.test(before)) return null;

      const fragment = parseTrailingWordFragment(before);
      if (!fragment) return null;
      const beforeFragment = before.slice(0, fragment.startCh);
      if (isImmediatelyAfterConfiguredPrefix(beforeFragment, plugin.getConfiguredPrefixes())) return null;

      const minChars = plugin.settings.proseSuggestionMinChars;
      const normalized = normalizeSearchText(fragment.text);
      if (normalized.length < minChars) return null;

      const records = plugin.getScopedEntitiesAllTypes(scope);
      const recordMatches = (record: EntityRecord, queryText: string, mode: "startsWith" | "includes"): boolean => {
        const matches = (candidate: string): boolean =>
          mode === "startsWith" ? candidate.startsWith(queryText) : candidate.includes(queryText);
        if (matches(record.normalizedName)) return true;
        return record.aliases.some((alias) => matches(normalizeSearchText(alias)));
      };

      const collectMatchingRecords = (queryText: string, mode: "startsWith" | "includes"): EntityRecord[] =>
        records.filter((record) => recordMatches(record, queryText, mode));

      const queryCandidates: string[] = [];
      for (let len = normalized.length; len >= minChars; len -= 1) {
        queryCandidates.push(normalized.slice(0, len));
      }

      let effectiveQuery = normalized;
      let matchedRecords: EntityRecord[] = [];
      for (const candidate of queryCandidates) {
        matchedRecords = collectMatchingRecords(candidate, "startsWith");
        if (matchedRecords.length > 0) {
          effectiveQuery = candidate;
          break;
        }
        matchedRecords = collectMatchingRecords(candidate, "includes");
        if (matchedRecords.length > 0) {
          effectiveQuery = candidate;
          break;
        }
      }
      if (matchedRecords.length === 0) return null;

      const suggestions: ProseMentionSuggestion[] = [];
      const seen = new Set<string>();
      for (const record of matchedRecords) {
        const canonicalNorm = record.normalizedName;
        const canonicalKey = `${record.file.path}::${canonicalNorm}`;
        if (!seen.has(canonicalKey)) {
          suggestions.push({
            record,
            displayText: record.name,
            normalizedDisplay: canonicalNorm,
            isAlias: false,
          });
          seen.add(canonicalKey);
        }

        for (const alias of record.aliases) {
          const aliasNorm = normalizeSearchText(alias);
          const aliasKey = `${record.file.path}::${aliasNorm}`;
          if (seen.has(aliasKey)) continue;
          suggestions.push({
            record,
            displayText: alias,
            normalizedDisplay: aliasNorm,
            isAlias: aliasNorm !== canonicalNorm,
          });
          seen.add(aliasKey);
        }
      }

      const rankedSuggestions = suggestions
        .sort(
          (a, b) =>
            scoreProseSuggestion(a, effectiveQuery) - scoreProseSuggestion(b, effectiveQuery) ||
            a.record.entityType.localeCompare(b.record.entityType) ||
            a.displayText.localeCompare(b.displayText),
        )
        .slice(0, 30);

      const options: Completion[] = rankedSuggestions.map((suggestion) => {
        const record = suggestion.record;
        const scopeLabel = record.scope === "shared" ? "Shared" : record.campaignId ?? "Campaign";
        const nameRole = suggestion.isAlias ? `alias of ${record.name}` : "canonical";
        const roleTag = suggestion.isAlias ? `A:${record.name}` : "C";
        return {
          label: `[${roleTag}] ${suggestion.displayText} (${record.entityType})`,
          detail: `${nameRole} | ${record.entityType} | ${scopeLabel}`,
          apply: (view: EditorView, _completion: unknown, from: number, to: number) => {
            const targetFile = plugin.getActiveFile();
            if (!targetFile) return;
            const link = plugin.app.metadataCache.fileToLinktext(record.file, targetFile.path, true);
            const typed = normalizeWhitespace(suggestion.displayText);
            const linkedText = suggestion.isAlias ? `[[${link}|${typed}]]` : `[[${link}]]`;
            view.dispatch({
              changes: { from, to, insert: linkedText },
              selection: { anchor: from + linkedText.length },
            });
          },
        };
      });

      return {
        from: line.from + fragment.startCh,
        options,
        filter: false,
      };
    };

    return [
      autocompletion({
        activateOnTyping: true,
        override: [source],
        icons: false,
        optionClass: (completion: Completion) => {
          const entityType = plugin.getEntityTypeFromCompletion(completion);
          if (!entityType) return "ttrpg-quicklog-completion-option";
          return `ttrpg-quicklog-completion-option ttrpg-quicklog-completion-type-${toCssSlug(entityType)}`;
        },
      }),
      EditorState.languageData.of(() => [{ autocomplete: source }]),
    ];
  }

  private getEntityTypeFromCompletion(completion: Completion): string | null {
    const detail = completion.detail;
    if (!detail) return null;
    const parts = detail.split("|").map((part) => part.trim());
    const entityType = parts.length >= 2 ? parts[1] : parts[0];
    if (!entityType) return null;
    return entityType;
  }

  private refreshCompletionColorStyles(): void {
    const knownTypes = this.getKnownEntityTypes();
    const rules: string[] = [];
    for (const type of knownTypes) {
      const color = this.getEffectiveColorForType(type);
      if (!color) continue;
      rules.push(
        `.cm-tooltip-autocomplete li.ttrpg-quicklog-completion-type-${toCssSlug(type)}::after { background: ${color}; }`,
      );
    }

    if (!this.completionColorStyleEl) {
      this.completionColorStyleEl = document.createElement("style");
      this.completionColorStyleEl.id = "ttrpg-quicklog-completion-colors";
      document.head.appendChild(this.completionColorStyleEl);
    }
    this.completionColorStyleEl.textContent = rules.join("\n");
  }

  private applyCategoryColorToElement(el: HTMLElement, color: string, type: string): void {
    el.classList.add("ttrpg-quicklog-colored-link");
    el.dataset.ttrpgType = type;
    el.style.setProperty("--ttrpg-link-color", color);
    el.style.setProperty("color", color, "important");
  }

  private clearCategoryColorFromElement(el: HTMLElement): void {
    el.classList.remove("ttrpg-quicklog-colored-link");
    delete el.dataset.ttrpgType;
    el.style.removeProperty("--ttrpg-link-color");
    el.style.removeProperty("color");
  }

  getConfiguredPrefixes(): string[] {
    return Array.from(new Set(this.settings.mentionTypes.map((m) => m.prefix)));
  }

  getKnownEntityTypes(): string[] {
    const fromMentions = this.settings.mentionTypes.map((m) => normalizeEntityType(m.entityType));
    const fromFolders = this.settings.campaigns.flatMap((c) => Object.keys(c.entityFolders || {}));
    return Array.from(new Set([...fromMentions, ...fromFolders].map(normalizeEntityType))).sort();
  }

  resolveScopeForFile(file: TFile): CampaignScope {
    const campaign = this.resolveCampaignForPath(file.path);
    const inShared = this.settings.sharedRoots.some((root) => isPathInScope(file.path, root));
    return { campaign, inShared };
  }

  resolveCampaignForPath(path: string): CampaignConfig | null {
    const normalized = normalizePath(path);
    const matching = this.settings.campaigns
      .filter((c) => isPathInScope(normalized, c.root))
      .sort((a, b) => b.root.length - a.root.length);
    return matching[0] ?? null;
  }

  getScopedEntities(scope: CampaignScope, entityType: string): EntityRecord[] {
    if (!scope.campaign) return [];

    const type = normalizeEntityType(entityType);
    const campaignEntities = this.entities.filter(
      (e) => e.scope === "campaign" && e.campaignId === scope.campaign?.id && e.entityType === type,
    );
    if (!this.settings.includeSharedInSuggestions) {
      return campaignEntities;
    }

    const sharedEntities = this.entities.filter((e) => e.scope === "shared" && e.entityType === type);
    return [...campaignEntities, ...sharedEntities];
  }

  getScopedEntitiesAllTypes(scope: CampaignScope): EntityRecord[] {
    if (!scope.campaign) return [];
    const campaignEntities = this.entities.filter((e) => e.scope === "campaign" && e.campaignId === scope.campaign?.id);
    if (!this.settings.includeSharedInSuggestions) {
      return campaignEntities;
    }
    const sharedEntities = this.entities.filter((e) => e.scope === "shared");
    return [...campaignEntities, ...sharedEntities];
  }

  async createEntityNote(
    name: string,
    entityType: string,
    campaign: CampaignConfig,
    options?: { skipVariantBackfill?: boolean },
  ): Promise<TFile | null> {
    const clean = sanitizeFileName(name);
    if (!clean) return null;

    const folderRelative =
      campaign.entityFolders[normalizeEntityType(entityType)] ??
      defaultFolderForType(entityType, this.settings.capitalizeDefaultCategoryFolders);
    const folderPath = normalizePath(`${campaign.root}/${folderRelative}`);
    await ensureFolder(this.app, folderPath);

    let targetPath = normalizePath(`${folderPath}/${clean}.md`);
    const adopted = await this.maybeAdoptExistingEntityNote(clean, entityType, campaign, targetPath);
    let entityFile: TFile | null = null;
    if (adopted) {
      entityFile = adopted;
    } else {
      if (this.app.vault.getAbstractFileByPath(targetPath)) {
        let i = 2;
        while (this.app.vault.getAbstractFileByPath(normalizePath(`${folderPath}/${clean} ${i}.md`))) {
          i += 1;
        }
        targetPath = normalizePath(`${folderPath}/${clean} ${i}.md`);
      }

      const file = await this.app.vault.create(targetPath, `# ${clean}\n`);
      await this.app.fileManager.processFrontMatter(file, (frontmatter) => {
        frontmatter.type = normalizeEntityType(entityType);
        frontmatter.campaign = campaign.id;
        if (!frontmatter.aliases) {
          frontmatter.aliases = [];
        }
      });
      entityFile = file;
    }

    if (!entityFile) return null;

    if (!options?.skipVariantBackfill) {
      await this.maybeOfferVariantBackfill(entityFile, clean, campaign);
    }
    await this.rebuildEntityIndex();
    return entityFile;
  }

  async offerVariantBackfill(entityFile: TFile, canonicalName: string, campaign: CampaignConfig): Promise<void> {
    await this.maybeOfferVariantBackfill(entityFile, canonicalName, campaign);
    await this.rebuildEntityIndex();
  }

  async updateEntityFromCampaignProse(entityFile: TFile, campaign: CampaignConfig): Promise<void> {
    const matches = await this.findVariantMatchesInCampaign(entityFile.basename, campaign.root, entityFile.path);
    if (matches.length === 0) {
      new Notice(`No variant spellings found for ${entityFile.basename}.`);
      return;
    }
    await this.offerVariantBackfill(entityFile, entityFile.basename, campaign);
  }

  private handlePrefixOverlayHotkey(evt: KeyboardEvent): void {
    const target = evt.target as HTMLElement | null;
    if (target && (target.tagName === "INPUT" || target.tagName === "TEXTAREA")) {
      return;
    }
    const configured = this.settings.prefixOverlayHotkey;
    if (!configured) return;
    if (!hotkeyMatchesEvent(configured, evt)) return;
    evt.preventDefault();
    this.togglePrefixOverlay();
  }

  togglePrefixOverlay(): void {
    if (this.prefixOverlayEl) {
      this.hidePrefixOverlay();
      return;
    }
    this.showPrefixOverlay();
  }

  private showPrefixOverlay(): void {
    const overlay = document.body.createDiv({ cls: "ttrpg-quicklog-prefix-overlay" });
    this.prefixOverlayEl = overlay;
    this.renderPrefixOverlay();
  }

  private hidePrefixOverlay(): void {
    if (!this.prefixOverlayEl) return;
    this.prefixOverlayEl.remove();
    this.prefixOverlayEl = null;
  }

  private renderPrefixOverlay(): void {
    const overlay = this.prefixOverlayEl;
    if (!overlay) return;
    overlay.empty();

    const panel = overlay.createDiv({ cls: "ttrpg-quicklog-prefix-overlay-panel" });
    panel.createEl("h3", { text: "Category Mention Reference" });
    panel.createEl("small", {
      text: `Hotkey: ${this.settings.prefixOverlayHotkey || "(disabled)"}`,
    });

    const list = panel.createEl("ul", { cls: "ttrpg-quicklog-prefix-overlay-list" });
    this.settings.mentionTypes.forEach((mapping) => {
      const item = list.createEl("li", { cls: "ttrpg-quicklog-prefix-overlay-item" });
      const swatch = item.createSpan({ cls: "ttrpg-quicklog-prefix-overlay-swatch" });
      const color = this.getEffectiveColorForType(mapping.entityType);
      if (color) {
        swatch.style.setProperty("background-color", color);
      }
      item.createSpan({ text: mapping.prefix, cls: "ttrpg-quicklog-prefix-overlay-prefix" });
      item.createSpan({ text: " -> ", cls: "ttrpg-quicklog-prefix-overlay-arrow" });
      item.createSpan({ text: mapping.entityType, cls: "ttrpg-quicklog-prefix-overlay-type" });
    });
  }

  private async maybeOfferVariantBackfill(
    entityFile: TFile,
    canonicalName: string,
    campaign: CampaignConfig,
    seedAliases: string[] = [],
  ): Promise<void> {
    const baseSeeds = Array.from(
      new Set([canonicalName, ...seedAliases].map((seed) => normalizeWhitespace(seed)).filter(Boolean)),
    );
    const matches = await this.findVariantMatchesForSeeds(baseSeeds, campaign.root, entityFile.path);
    const canonicalNorm = normalizeSearchText(canonicalName);
    if (matches.length === 0) return;

    const decision = await VariantBackfillModal.prompt(
      this.app,
      canonicalName,
      matches,
      async (customSpellings) =>
        this.findVariantMatchesForSeeds([...baseSeeds, ...customSpellings], campaign.root, entityFile.path),
    );
    if (!decision.applyLinks && !decision.addAliases) return;
    const selectedVariants = Array.from(
      new Set(decision.selectedVariants.map((v) => normalizeWhitespace(v)).filter(Boolean)),
    );
    if (selectedVariants.length === 0) return;

    if (decision.addAliases) {
      const aliasVariants = selectedVariants.filter(
        (variant) => normalizeSearchText(variant) !== canonicalNorm,
      );
      const addedAliases = await this.addVariantsAsAliases(entityFile, canonicalName, aliasVariants);
      if (addedAliases > 0) {
        new Notice(`Added ${addedAliases} alias variant(s) to ${entityFile.basename}.`);
      }
    }

    if (decision.applyLinks) {
      const replacements = await this.applyVariantLinksInCampaign(
        entityFile,
        campaign.root,
        selectedVariants,
      );
      if (replacements > 0) {
        new Notice(`Linked ${replacements} occurrence(s) to ${entityFile.basename}.`);
      }
    }
  }

  private async addVariantsAsAliases(entityFile: TFile, canonicalName: string, variants: string[]): Promise<number> {
    let added = 0;
    const canonicalNorm = normalizeSearchText(canonicalName);
    const uniqueVariants = Array.from(new Set(variants.map((v) => normalizeWhitespace(v)).filter(Boolean)));

    await this.app.fileManager.processFrontMatter(entityFile, (frontmatter) => {
      const existing = frontmatterAliasesToArray(frontmatter.aliases);
      const existingNorm = new Set(existing.map((v) => normalizeSearchText(v)));
      const next = [...existing];

      for (const variant of uniqueVariants) {
        const variantNorm = normalizeSearchText(variant);
        if (variantNorm === canonicalNorm) continue;
        if (existingNorm.has(variantNorm)) continue;
        next.push(variant);
        existingNorm.add(variantNorm);
        added += 1;
      }

      frontmatter.aliases = next;
    });

    return added;
  }

  private async applyVariantLinksInCampaign(
    entityFile: TFile,
    campaignRoot: string,
    variants: string[],
  ): Promise<number> {
    const sortedVariants = Array.from(new Set(variants.map((v) => normalizeWhitespace(v)).filter(Boolean))).sort(
      (a, b) => b.length - a.length,
    );
    if (sortedVariants.length === 0) return 0;

    let replacements = 0;
    const files = this.app.vault
      .getMarkdownFiles()
      .filter((file) => isPathInScope(file.path, campaignRoot) && file.path !== entityFile.path);

    for (const file of files) {
      const original = await this.app.vault.cachedRead(file);
      const linkPath = this.app.metadataCache.fileToLinktext(entityFile, file.path, true);
      const updated = replaceVariantOccurrencesInText(original, sortedVariants, linkPath, () => {
        replacements += 1;
      });
      if (updated !== original) {
        await this.app.vault.modify(file, updated);
      }
    }

    return replacements;
  }

  private async maybeAdoptExistingEntityNote(
    cleanName: string,
    entityType: string,
    campaign: CampaignConfig,
    preferredPath: string,
  ): Promise<TFile | null> {
    const matching = this.app.vault
      .getMarkdownFiles()
      .filter(
        (file) =>
          isPathInScope(file.path, campaign.root) &&
          normalizeSearchText(file.basename) === normalizeSearchText(cleanName),
      )
      .sort((a, b) => {
        if (a.path === preferredPath) return -1;
        if (b.path === preferredPath) return 1;
        return a.path.length - b.path.length;
      });

    const candidate = matching[0];
    if (!candidate) return null;

    const adopt = await AdoptExistingEntityModal.prompt(this.app, {
      name: cleanName,
      path: candidate.path,
    });
    if (!adopt) return null;

    await this.app.fileManager.processFrontMatter(candidate, (frontmatter) => {
      if (frontmatter.type === undefined) {
        frontmatter.type = normalizeEntityType(entityType);
      }
      if (frontmatter.campaign === undefined) {
        frontmatter.campaign = campaign.id;
      }
      if (frontmatter.aliases === undefined) {
        frontmatter.aliases = [cleanName];
      } else if (typeof frontmatter.aliases === "string") {
        const existing = frontmatter.aliases.trim();
        if (!existing) {
          frontmatter.aliases = [cleanName];
        } else if (normalizeSearchText(existing) !== normalizeSearchText(cleanName)) {
          frontmatter.aliases = [existing, cleanName];
        }
      } else if (Array.isArray(frontmatter.aliases)) {
        const hasAlias = frontmatter.aliases.some(
          (alias: unknown) => normalizeSearchText(String(alias)) === normalizeSearchText(cleanName),
        );
        if (!hasAlias) {
          frontmatter.aliases.push(cleanName);
        }
      }
    });

    return candidate;
  }

  private async findVariantMatchesInCampaign(
    canonicalName: string,
    campaignRoot: string,
    excludePath: string,
  ): Promise<VariantMatch[]> {
    const target = normalizeFuzzyText(canonicalName);
    const targetWords = splitWords(target);
    if (!target || targetWords.length === 0) return [];

    const files = this.app.vault
      .getMarkdownFiles()
      .filter((file) => isPathInScope(file.path, campaignRoot) && file.path !== excludePath);

    const byVariant = new Map<string, VariantMatch>();

    for (const file of files) {
      const content = await this.app.vault.cachedRead(file);
      for (const proseLine of iterateProseLines(content)) {
        const candidates = collectFuzzyCandidatesFromLine(proseLine, targetWords.length);
        for (const candidate of candidates) {
          const normalizedCandidate = normalizeFuzzyText(candidate);
          if (!normalizedCandidate) continue;
          if (!looksLikeSameEntity(target, normalizedCandidate)) continue;
          const threshold = maxLevenshteinFor(normalizedCandidate);
          const distance = levenshtein(normalizedCandidate, target);
          if (distance > threshold) continue;

          const key = normalizeWhitespace(candidate);
          if (!key) continue;
          const existing = byVariant.get(key);
          if (existing) {
            existing.count += 1;
            existing.files.add(file.path);
          } else {
            byVariant.set(key, {
              variant: key,
              count: 1,
              files: new Set([file.path]),
            });
          }
        }
      }
    }

    return Array.from(byVariant.values()).sort((a, b) => b.count - a.count || a.variant.localeCompare(b.variant));
  }

  private async findVariantMatchesForSeeds(
    seedNames: string[],
    campaignRoot: string,
    excludePath: string,
  ): Promise<VariantMatch[]> {
    const uniqueSeeds = Array.from(
      new Set(seedNames.map((seed) => normalizeWhitespace(seed)).filter((seed) => seed.length > 0)),
    );
    if (uniqueSeeds.length === 0) return [];

    const merged = new Map<string, VariantMatch>();
    for (const seed of uniqueSeeds) {
      const matches = await this.findVariantMatchesInCampaign(seed, campaignRoot, excludePath);
      for (const match of matches) {
        const key = normalizeWhitespace(match.variant);
        if (!key) continue;
        const existing = merged.get(key);
        if (!existing) {
          merged.set(key, {
            variant: key,
            count: match.count,
            files: new Set(match.files),
          });
          continue;
        }
        existing.count += match.count;
        for (const path of match.files) {
          existing.files.add(path);
        }
      }
    }

    return Array.from(merged.values()).sort((a, b) => b.count - a.count || a.variant.localeCompare(b.variant));
  }

  convertMentionsInText(text: string, activeFile: TFile, scope: CampaignScope): string {
    const prefixes = this.getConfiguredPrefixes();
    if (!scope.campaign || prefixes.length === 0) return text;

    const alts = [...prefixes]
      .sort((a, b) => b.length - a.length)
      .map(escapeRegex)
      .join("|");
    const modes = this.settings.mentionInputModes;
    const unquotedPattern = modes.open
      ? "([^\\n\\r.,;:!?()[\\]{}]+)"
      : "([^\\s\\[\\]]+)";
    const pattern = new RegExp(`(^|\\s)(${alts})(?:"([^"]+)"|${unquotedPattern})`, "g");

    return text.replace(pattern, (full, leading, prefix, quoted, plain) => {
      const mentionType = this.getMentionTypeForPrefix(prefix);
      if (!mentionType) return full;

      const raw = (quoted || plain || "").trim();
      if (!raw) return full;
      const query = normalizeSearchText(raw);
      const candidates = this.getScopedEntities(scope, mentionType.entityType);
      const exact = candidates.find(
        (record) => record.normalizedName === query || record.normalizedAliases.includes(query),
      );
      if (!exact) return full;

      const link = this.app.metadataCache.fileToLinktext(exact.file, activeFile.path, true);
      return `${leading}[[${link}]]`;
    });
  }

  getMaintenanceScopedMarkdownFiles(): TFile[] {
    const all = this.app.vault.getMarkdownFiles();
    const roots = [
      ...this.settings.campaigns.map((c) => normalizePath(c.root)),
      ...this.settings.sharedRoots.map((r) => normalizePath(r)),
    ];
    if (roots.length === 0) return all;
    return all.filter((file) => roots.some((root) => isPathInScope(file.path, root)));
  }

  findFilesWithCategoryType(entityType: string): TFile[] {
    const normalizedType = normalizeEntityType(entityType);
    return this.getMaintenanceScopedMarkdownFiles().filter((file) => {
      const fm = this.app.metadataCache.getFileCache(file)?.frontmatter;
      const typeRaw = fm?.type;
      if (typeof typeRaw !== "string") return false;
      return normalizeEntityType(typeRaw) === normalizedType;
    });
  }

  async replaceCategoryTypeInFiles(files: TFile[], oldType: string, newType: string): Promise<number> {
    const oldNormalized = normalizeEntityType(oldType);
    const newNormalized = normalizeEntityType(newType);
    let changed = 0;

    for (const file of files) {
      await this.app.fileManager.processFrontMatter(file, (frontmatter) => {
        const raw = frontmatter.type;
        if (typeof raw !== "string") return;
        if (normalizeEntityType(raw) !== oldNormalized) return;
        frontmatter.type = newNormalized;
        changed += 1;
      });
    }
    await this.rebuildEntityIndex();
    return changed;
  }

  async renameCategoryFoldersOnDisk(oldType: string, newType: string): Promise<number> {
    const newFolderName = formatCategoryFolderName(newType, this.settings.capitalizeDefaultCategoryFolders);
    let renamed = 0;

    for (const campaign of this.settings.campaigns) {
      const oldRelative = campaign.entityFolders[oldType];
      if (!oldRelative) continue;

      const normalizedOldRelative = normalizePath(oldRelative);
      const parentParts = normalizedOldRelative.split("/");
      parentParts.pop();
      const parent = parentParts.join("/");

      const newRelative = normalizePath(parent ? `${parent}/${newFolderName}` : newFolderName);
      const oldAbs = normalizePath(`${campaign.root}/${normalizedOldRelative}`);
      const newAbs = normalizePath(`${campaign.root}/${newRelative}`);
      const oldFolder = this.app.vault.getAbstractFileByPath(oldAbs);
      const newExists = this.app.vault.getAbstractFileByPath(newAbs);

      if (oldFolder instanceof TFolder && !newExists) {
        await this.app.vault.rename(oldFolder, newAbs);
        renamed += 1;
      }
      campaign.entityFolders[newType] = newRelative;
      delete campaign.entityFolders[oldType];
    }

    return renamed;
  }

  getCategoryTypesForDefaults(): string[] {
    return Array.from(
      new Set(this.settings.mentionTypes.map((m) => normalizeEntityType(m.entityType)).filter(Boolean)),
    );
  }

  private inferEntityTypeFromCampaignFolder(filePath: string, campaign: CampaignConfig): string | null {
    const types = this.getKnownEntityTypes();
    for (const type of types) {
      const folderRelative =
        campaign.entityFolders[type] ?? defaultFolderForType(type, this.settings.capitalizeDefaultCategoryFolders);
      const folderAbsolute = normalizePath(`${campaign.root}/${folderRelative}`);
      if (isPathInScope(filePath, folderAbsolute)) {
        return normalizeEntityType(type);
      }
    }
    return null;
  }

  private inferEntityTypeFromSharedFolder(filePath: string): string | null {
    const types = this.getKnownEntityTypes();
    for (const root of this.settings.sharedRoots) {
      for (const type of types) {
        const folderRelative = defaultFolderForType(type, this.settings.capitalizeDefaultCategoryFolders);
        const folderAbsolute = normalizePath(`${root}/${folderRelative}`);
        if (isPathInScope(filePath, folderAbsolute)) {
          return normalizeEntityType(type);
        }
      }
    }
    return null;
  }

  getMissingCategoryFolderPathsForCampaign(campaign: CampaignConfig): string[] {
    const missing: string[] = [];
    const knownTypes = this.getCategoryTypesForDefaults();

    for (const categoryType of knownTypes) {
      const relative =
        campaign.entityFolders[categoryType] ??
        defaultFolderForType(categoryType, this.settings.capitalizeDefaultCategoryFolders);
      const absolute = normalizePath(`${campaign.root}/${relative}`);
      if (!this.app.vault.getAbstractFileByPath(absolute)) {
        missing.push(absolute);
      }
    }

    return missing;
  }

  async createMissingCategoryFoldersForCampaign(campaign: CampaignConfig): Promise<number> {
    const missing = this.getMissingCategoryFolderPathsForCampaign(campaign);
    for (const path of missing) {
      await ensureFolder(this.app, path);
    }
    return missing.length;
  }

  async offerCreateCategoryFoldersForCampaign(campaign: CampaignConfig): Promise<void> {
    const missing = this.getMissingCategoryFolderPathsForCampaign(campaign);
    if (missing.length === 0) return;

    const shouldCreate = await CreateCampaignFoldersModal.prompt(this.app, missing.length, campaign.name);
    if (!shouldCreate) return;

    const createdCount = await this.createMissingCategoryFoldersForCampaign(campaign);
    if (createdCount > 0) {
      new Notice(`Created ${createdCount} category folder(s) for ${campaign.name}.`);
    }
  }

  private async rebuildEntityIndex(): Promise<void> {
    const files = this.app.vault.getMarkdownFiles();
    const next: EntityRecord[] = [];

    for (const file of files) {
      const scope = this.resolveScopeForFile(file);
      if (!scope.campaign && !scope.inShared) continue;

      const cache = this.app.metadataCache.getFileCache(file);
      const fm = cache?.frontmatter;
      const aliases = frontmatterAliasesToArray(fm?.aliases);
      const typeFromFrontmatter = typeof fm?.type === "string" ? normalizeEntityType(fm.type) : "";
      const typeFromFolder = scope.campaign
        ? this.inferEntityTypeFromCampaignFolder(file.path, scope.campaign)
        : scope.inShared
          ? this.inferEntityTypeFromSharedFolder(file.path)
          : null;
      const isLikelyEntity = Boolean(typeFromFrontmatter || aliases.length > 0 || typeFromFolder);
      if (!isLikelyEntity) continue;

      const type = typeFromFrontmatter || typeFromFolder || "entity";
      const name = file.basename;

      next.push({
        file,
        name,
        normalizedName: normalizeSearchText(name),
        aliases,
        normalizedAliases: aliases.map(normalizeSearchText),
        entityType: type,
        scope: scope.campaign ? "campaign" : "shared",
        campaignId: scope.campaign?.id ?? null,
      });
    }

    this.entities = next;
  }
}

class CampaignMentionSuggester extends EditorSuggest<MentionSuggestion> {
  plugin: CampaignNamespaceLinkerPlugin;

  constructor(plugin: CampaignNamespaceLinkerPlugin) {
    super(plugin.app);
    this.plugin = plugin;
  }

  onTrigger(
    cursor: EditorPosition,
    editor: Editor,
    _file: TFile | null,
  ): EditorSuggestTriggerInfo | null {
    const line = editor.getLine(cursor.line).slice(0, cursor.ch);
    const parsed = parseMentionAtEnd(
      line,
      this.plugin.getConfiguredPrefixes(),
      this.plugin.settings.mentionInputModes,
    );
    if (!parsed) return null;

    const mentionType = this.plugin.getMentionTypeForPrefix(parsed.prefix);
    if (!mentionType) return null;

    const query = JSON.stringify({
      prefix: parsed.prefix,
      rawQuery: parsed.rawQuery,
      entityType: mentionType.entityType,
    });

    return {
      start: { line: cursor.line, ch: parsed.startCh },
      end: cursor,
      query,
    };
  }

  getSuggestions(context: EditorSuggestContext): MentionSuggestion[] {
    const parsed = parseMentionQuery(context.query);
    if (!parsed) return [];

    const file = this.plugin.getActiveFile();
    if (!file) return [];
    const scope = this.plugin.resolveScopeForFile(file);
    if (!scope.campaign) return [];

    const rawForCreate = normalizeWhitespace(parsed.rawQuery.replace(/^"|"$/g, ""));
    const query = normalizeSearchText(rawForCreate);
    const records = this.plugin.getScopedEntities(scope, parsed.entityType);
    const filtered = records
      .filter((record) => {
        if (!query) return true;
        return (
          record.normalizedName.includes(query) ||
          record.normalizedAliases.some((alias) => alias.includes(query))
        );
      })
      .sort((a, b) => scoreRecord(a, query) - scoreRecord(b, query) || a.name.localeCompare(b.name))
      .slice(0, 30);

    const suggestions: MentionSuggestion[] = filtered.map((record) => ({
      kind: "existing",
      record,
    }));

    if (query) {
      const exactMatch = records.some(
        (record) => record.normalizedName === query || record.normalizedAliases.includes(query),
      );
      if (!exactMatch && scope.campaign) {
        suggestions.unshift({
          kind: "create",
          name: rawForCreate,
          entityType: parsed.entityType,
          campaign: scope.campaign,
        });
      }
    }

    return suggestions;
  }

  renderSuggestion(suggestion: MentionSuggestion, el: HTMLElement): void {
    if (suggestion.kind === "create") {
      el.createEl("div", { text: `Create ${suggestion.entityType}: ${suggestion.name}` });
      el.createEl("small", { text: `Campaign: ${suggestion.campaign.name}` });
      return;
    }

    const record = suggestion.record;
    el.createEl("div", { text: record.name });
    const scopeLabel = record.scope === "shared" ? "Shared" : record.campaignId ?? "Campaign";
    el.createEl("small", { text: `${record.entityType} | ${scopeLabel}` });
  }

  selectSuggestion(suggestion: MentionSuggestion, evt: MouseEvent | KeyboardEvent): void {
    const context = this.context;
    if (!context) return;

    const view = this.app.workspace.getActiveViewOfType(MarkdownView);
    if (!view) return;

    const editor = view.editor;
    const file = view.file;
    if (!file) return;

    const writeLink = (linkText: string): void => {
      editor.replaceRange(`[[${linkText}]]`, context.start, context.end);
    };

    if (suggestion.kind === "existing") {
      const link = this.app.metadataCache.fileToLinktext(suggestion.record.file, file.path, true);
      writeLink(link);
      this.close();
      return;
    }

    void (async () => {
      const created = await this.plugin.createEntityNote(
        suggestion.name,
        suggestion.entityType,
        suggestion.campaign,
        { skipVariantBackfill: true },
      );
      if (!created) {
        new Notice("Could not create entity note.");
        return;
      }
      const link = this.app.metadataCache.fileToLinktext(created, file.path, true);
      writeLink(link);
      await this.plugin.offerVariantBackfill(created, created.basename, suggestion.campaign);
      new Notice(`Created ${suggestion.entityType}: ${created.basename}`);
      this.close();
    })();

    if (evt instanceof KeyboardEvent) {
      evt.preventDefault();
    }
  }
}

class CampaignNamespaceLinkerSettingTab extends PluginSettingTab {
  plugin: CampaignNamespaceLinkerPlugin;

  constructor(app: App, plugin: CampaignNamespaceLinkerPlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();

    containerEl.createEl("h2", { text: "TTRPG QuickLog" });

    new Setting(containerEl)
      .setName("Include shared roots in suggestions")
      .setDesc("When enabled, campaign-local suggestions also include matching entities from shared roots.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.settings.includeSharedInSuggestions).onChange(async (value) => {
          this.plugin.settings.includeSharedInSuggestions = value;
          await this.plugin.saveSettings();
        }),
      );

    new Setting(containerEl)
      .setName("Enable prefix-less prose suggestions")
      .setDesc("When enabled, typing prose without prefixes can suggest existing entities and aliases.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.settings.enableProseSuggestions).onChange(async (value) => {
          this.plugin.settings.enableProseSuggestions = value;
          await this.plugin.saveSettings();
        }),
      );

    new Setting(containerEl)
      .setName("Minimum prose suggestion characters")
      .setDesc("Minimum number of typed characters before prefix-less suggestions appear.")
      .addText((text) => {
        text
          .setPlaceholder("2")
          .setValue(String(this.plugin.settings.proseSuggestionMinChars))
          .onChange(async (value) => {
            const parsed = Number.parseInt(value.trim(), 10);
            this.plugin.settings.proseSuggestionMinChars = clampNumber(
              parsed,
              1,
              10,
              DEFAULT_SETTINGS.proseSuggestionMinChars,
            );
            await this.plugin.saveSettings();
          });
      });

    new Setting(containerEl)
      .setName("Prefix reference overlay hotkey")
      .setDesc("Click the field and press a modifier combo (for example Ctrl+Shift+O). Backspace/Delete clears.")
      .addText((text) => {
        text.setPlaceholder("Mod+Shift+O").setValue(this.plugin.settings.prefixOverlayHotkey);
        text.inputEl.addEventListener("keydown", async (evt) => {
          if (evt.key === "Backspace" || evt.key === "Delete") {
            evt.preventDefault();
            this.plugin.settings.prefixOverlayHotkey = "";
            text.setValue("");
            await this.plugin.saveSettings();
            return;
          }

          const combo = formatHotkeyFromEvent(evt);
          if (!combo) return;
          evt.preventDefault();
          this.plugin.settings.prefixOverlayHotkey = combo;
          text.setValue(combo);
          await this.plugin.saveSettings();
        });
      });

    this.renderMentionTypes(containerEl);
    this.renderMentionInputModes(containerEl);
    this.renderCampaigns(containerEl);
    this.renderSharedRoots(containerEl);
  }

  private renderCampaigns(containerEl: HTMLElement): void {
    containerEl.createEl("h3", { text: "Campaign roots" });

    new Setting(containerEl)
      .setName("Capitalize default category folders")
      .setDesc(
        "When enabled, autogenerated default folder names are title-cased. Custom folder paths are not changed.",
      )
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.settings.capitalizeDefaultCategoryFolders).onChange(async (value) => {
          this.plugin.settings.capitalizeDefaultCategoryFolders = value;
          await this.plugin.saveSettings();
          this.display();
        }),
      );

    for (const campaign of this.plugin.settings.campaigns) {
      const box = containerEl.createDiv("cnl-setting-box");
      box.createEl("h4", { text: campaign.name });

      new Setting(box).setName("Campaign id").addText((text) =>
        text.setValue(campaign.id).onChange(async (value) => {
          campaign.id = slugify(value.trim() || campaign.name);
          await this.plugin.saveSettings();
        }),
      );

      new Setting(box).setName("Campaign name").addText((text) =>
        text.setValue(campaign.name).onChange(async (value) => {
          campaign.name = value.trim();
          if (!campaign.name) campaign.name = campaign.id;
          await this.plugin.saveSettings();
        }),
      );

      new Setting(box).setName("Root folder").addText((text) =>
        text.setValue(campaign.root).onChange(async (value) => {
          campaign.root = normalizePath(value.trim());
          await this.plugin.saveSettings();
        }),
      );

      box.createEl("div", { text: "Category Entity Folders (relative to campaign root)" });
      for (const entityType of this.plugin.getKnownEntityTypes()) {
        new Setting(box).setName(entityType).addText((text) =>
          text
            .setPlaceholder(
              defaultFolderForType(entityType, this.plugin.settings.capitalizeDefaultCategoryFolders),
            )
            .setValue(campaign.entityFolders[entityType] ?? "")
            .onChange(async (value) => {
              const clean = normalizePath(value.trim());
              if (!clean) delete campaign.entityFolders[entityType];
              else campaign.entityFolders[entityType] = clean;
              await this.plugin.saveSettings();
            }),
        );
      }

      new Setting(box).addButton((btn) =>
        btn
          .setWarning()
          .setButtonText("Remove campaign")
          .onClick(async () => {
            this.plugin.settings.campaigns = this.plugin.settings.campaigns.filter((c) => c !== campaign);
            await this.plugin.saveSettings();
            this.display();
          }),
      );
    }

    let newName = "";
    let newRoot = "";
    new Setting(containerEl)
      .setName("Add campaign")
      .addText((text) => text.setPlaceholder("Emberfall").onChange((value) => (newName = value.trim())))
      .addText((text) =>
        text
          .setPlaceholder("Campaigns/Emberfall")
          .onChange((value) => (newRoot = normalizePath(value.trim()))),
      )
      .addButton((btn) =>
        btn.setButtonText("Add").onClick(async () => {
          if (!newName || !newRoot) {
            new Notice("Campaign name and root are required.");
            return;
          }

          if (this.plugin.settings.campaigns.some((c) => c.root === newRoot)) {
            new Notice("Campaign root already exists.");
            return;
          }

          const campaign: CampaignConfig = {
            id: slugify(newName),
            name: newName,
            root: newRoot,
            entityFolders: {},
          };
          this.plugin.settings.campaigns.push(campaign);
          await this.plugin.saveSettings();
          await this.plugin.offerCreateCategoryFoldersForCampaign(campaign);
          this.display();
        }),
      );
  }

  private renderSharedRoots(containerEl: HTMLElement): void {
    containerEl.createEl("h3", { text: "Shared roots" });

    this.plugin.settings.sharedRoots.forEach((root, index) => {
      new Setting(containerEl)
        .setName(`Shared root ${index + 1}`)
        .addText((text) =>
          text.setValue(root).onChange(async (value) => {
            this.plugin.settings.sharedRoots[index] = normalizePath(value.trim());
            await this.plugin.saveSettings();
          }),
        )
        .addButton((btn) =>
          btn
            .setWarning()
            .setButtonText("Remove")
            .onClick(async () => {
              this.plugin.settings.sharedRoots.splice(index, 1);
              await this.plugin.saveSettings();
              this.display();
            }),
        );
    });

    let newSharedRoot = "";
    new Setting(containerEl)
      .setName("Add shared root")
      .addText((text) =>
        text.setPlaceholder("Shared").onChange((value) => {
          newSharedRoot = normalizePath(value.trim());
        }),
      )
      .addButton((btn) =>
        btn.setButtonText("Add").onClick(async () => {
          if (!newSharedRoot) {
            new Notice("Provide a shared root path.");
            return;
          }
          if (this.plugin.settings.sharedRoots.includes(newSharedRoot)) {
            new Notice("Shared root already exists.");
            return;
          }
          this.plugin.settings.sharedRoots.push(newSharedRoot);
          await this.plugin.saveSettings();
          this.display();
        }),
      );
  }

  private renderMentionTypes(containerEl: HTMLElement): void {
    containerEl.createEl("h3", { text: "Category Mention Prefixes" });
    containerEl.createEl("p", {
      text: "Map each prefix to a category entity type. Example: @@ -> npcs, ## -> locations",
    });
    new Setting(containerEl)
      .setName("Use rainbow fallback colors")
      .setDesc("If a category has no custom color, assign one from rainbow order top-to-bottom.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.settings.useRainbowCategoryColors).onChange(async (value) => {
          this.plugin.settings.useRainbowCategoryColors = value;
          await this.plugin.saveSettings();
          this.plugin.queueLinkColorRefresh();
        }),
      );

    this.plugin.settings.mentionTypes.forEach((mapping, index) => {
      new Setting(containerEl)
        .setName(`Category Mention ${index + 1}`)
        .addText((text) =>
          text.setValue(mapping.prefix).onChange(async (value) => {
            mapping.prefix = value;
            await this.plugin.saveSettings();
          }),
        )
        .addText((text) => {
          const originalType = normalizeEntityType(mapping.entityType);
          let pendingType = originalType;

          text.setValue(mapping.entityType).onChange((value) => {
            pendingType = normalizeEntityType(value);
            mapping.entityType = pendingType;
          });

          text.inputEl.addEventListener("blur", async () => {
            if (!pendingType) {
              mapping.entityType = originalType;
              text.setValue(originalType);
              new Notice("Category type cannot be empty.");
              return;
            }

            if (pendingType !== originalType) {
              const oldTypeStillUsed = this.plugin.settings.mentionTypes.some(
                (m, idx) => idx !== index && normalizeEntityType(m.entityType) === originalType,
              );

              if (!oldTypeStillUsed) {
                const filesWithOldType = this.plugin.findFilesWithCategoryType(originalType);
                const hasFolderMappings = this.plugin.settings.campaigns.some(
                  (campaign) => campaign.entityFolders[originalType] !== undefined,
                );

                let decision: CategoryRenameDecision = {
                  applyTypeUpdate: false,
                  renameFolders: false,
                };

                if (filesWithOldType.length > 0 || hasFolderMappings) {
                  decision = await CategoryRenameModal.prompt(this.app, {
                    oldType: originalType,
                    newType: pendingType,
                    affectedNotes: filesWithOldType.length,
                    hasFolderMappings,
                    capitalizeDefaultCategoryFolders: this.plugin.settings.capitalizeDefaultCategoryFolders,
                  });
                }

                if (decision.renameFolders && hasFolderMappings) {
                  const renamedCount = await this.plugin.renameCategoryFoldersOnDisk(
                    originalType,
                    pendingType,
                  );
                  if (renamedCount > 0) {
                    new Notice(`Renamed ${renamedCount} category folder(s) on disk.`);
                  } else {
                    new Notice("No folders renamed (destination exists or source missing).");
                  }
                } else {
                  this.remapCategoryEntityFolderKeys(originalType, pendingType);
                }

                if (decision.applyTypeUpdate && filesWithOldType.length > 0) {
                  const updated = await this.plugin.replaceCategoryTypeInFiles(
                    filesWithOldType,
                    originalType,
                    pendingType,
                  );
                  new Notice(`Updated ${updated} note type value(s) from ${originalType} to ${pendingType}.`);
                }
              }
            }

            await this.plugin.saveSettings();
            this.display();
          });
        })
        .addColorPicker((picker) =>
          picker
            .setValue(this.plugin.getEffectiveColorForType(mapping.entityType) ?? "#888888")
            .onChange(async (value) => {
              mapping.color = normalizeHexColor(value);
              await this.plugin.saveSettings();
              this.plugin.queueLinkColorRefresh();
            }),
        )
        .addButton((btn) =>
          btn
            .setWarning()
            .setButtonText("Remove")
            .onClick(async () => {
              this.plugin.settings.mentionTypes.splice(index, 1);
              await this.plugin.saveSettings();
              this.display();
            }),
        );
    });

    let newPrefix = "";
    let newEntityType = "";
    new Setting(containerEl)
      .setName("Category Mention Mapping")
      .addText((text) =>
        text.setPlaceholder("@").onChange((value) => {
          newPrefix = value.trim();
        }),
      )
      .addText((text) =>
        text.setPlaceholder("npc").onChange((value) => {
          newEntityType = normalizeEntityType(value.trim());
        }),
      )
      .addButton((btn) =>
        btn.setButtonText("Add").onClick(async () => {
          if (!newPrefix || !newEntityType) {
            new Notice("Prefix and entity type are required.");
            return;
          }
          this.plugin.settings.mentionTypes.push({
            prefix: newPrefix,
            entityType: newEntityType,
          });
          await this.plugin.saveSettings();
          this.display();
        }),
      );
  }

  private remapCategoryEntityFolderKeys(oldType: string, newType: string): void {
    for (const campaign of this.plugin.settings.campaigns) {
      const oldPath = campaign.entityFolders[oldType];
      if (oldPath === undefined) continue;

      if (!campaign.entityFolders[newType]) {
        campaign.entityFolders[newType] = oldPath;
      }
      delete campaign.entityFolders[oldType];
    }
  }

  private renderMentionInputModes(containerEl: HTMLElement): void {
    containerEl.createEl("h3", { text: "Mention input modes" });
    containerEl.createEl("p", {
      text: "Enable any combination. Enter confirms the selected suggestion by default.",
    });

    new Setting(containerEl)
      .setName("Single-word mode")
      .setDesc("Supports mentions like @Arlen")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.settings.mentionInputModes.singleWord).onChange(async (value) => {
          this.plugin.settings.mentionInputModes.singleWord = value;
          await this.plugin.saveSettings();
        }),
      );

    new Setting(containerEl)
      .setName("Quoted mode")
      .setDesc('Supports mentions like @"Captain Arlen"')
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.settings.mentionInputModes.quoted).onChange(async (value) => {
          this.plugin.settings.mentionInputModes.quoted = value;
          await this.plugin.saveSettings();
        }),
      );

    new Setting(containerEl)
      .setName("Open mode")
      .setDesc("Supports mentions like @Captain Arlen")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.settings.mentionInputModes.open).onChange(async (value) => {
          this.plugin.settings.mentionInputModes.open = value;
          await this.plugin.saveSettings();
        }),
      );
  }
}

class CategoryRenameModal extends Modal {
  private readonly oldType: string;
  private readonly newType: string;
  private readonly affectedNotes: number;
  private readonly hasFolderMappings: boolean;
  private readonly capitalizeDefaultCategoryFolders: boolean;
  private resolver: ((value: CategoryRenameDecision) => void) | null = null;

  private applyTypeUpdate = true;
  private renameFolders = false;

  constructor(
    app: App,
    data: {
      oldType: string;
      newType: string;
      affectedNotes: number;
      hasFolderMappings: boolean;
      capitalizeDefaultCategoryFolders: boolean;
    },
  ) {
    super(app);
    this.oldType = data.oldType;
    this.newType = data.newType;
    this.affectedNotes = data.affectedNotes;
    this.hasFolderMappings = data.hasFolderMappings;
    this.capitalizeDefaultCategoryFolders = data.capitalizeDefaultCategoryFolders;
    this.setTitle("Category Rename Follow-up");
  }

  static async prompt(
    app: App,
    data: {
      oldType: string;
      newType: string;
      affectedNotes: number;
      hasFolderMappings: boolean;
      capitalizeDefaultCategoryFolders: boolean;
    },
  ): Promise<CategoryRenameDecision> {
    const modal = new CategoryRenameModal(app, data);
    return modal.openAndWait();
  }

  private openAndWait(): Promise<CategoryRenameDecision> {
    return new Promise((resolve) => {
      this.resolver = resolve;
      this.open();
    });
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();

    contentEl.createEl("p", {
      text: `Category type changed from "${this.oldType}" to "${this.newType}".`,
    });
    contentEl.createEl("p", {
      text: `${this.affectedNotes} note(s) still use frontmatter type "${this.oldType}".`,
    });

    new Setting(contentEl)
      .setName("Update note type values")
      .setDesc(`Replace frontmatter type "${this.oldType}" with "${this.newType}" in affected notes.`)
      .addToggle((toggle) =>
        toggle.setValue(this.applyTypeUpdate).onChange((value) => {
          this.applyTypeUpdate = value;
        }),
      );

    new Setting(contentEl)
      .setName("Rename category folders on disk")
      .setDesc(
        `Rename mapped folders to "${formatCategoryFolderName(this.newType, this.capitalizeDefaultCategoryFolders)}" where possible.`,
      )
      .addToggle((toggle) =>
        toggle.setValue(this.renameFolders).setDisabled(!this.hasFolderMappings).onChange((value) => {
          this.renameFolders = value;
        }),
      );

    if (!this.hasFolderMappings) {
      contentEl.createEl("small", {
        text: "No mapped category folders found for the old category type.",
      });
    }

    new Setting(contentEl)
      .addButton((btn) =>
        btn.setButtonText("Apply selected changes").setCta().onClick(() => {
          this.finish({
            applyTypeUpdate: this.applyTypeUpdate,
            renameFolders: this.renameFolders && this.hasFolderMappings,
          });
        }),
      )
      .addButton((btn) =>
        btn.setButtonText("Leave as-is").onClick(() => {
          this.finish({
            applyTypeUpdate: false,
            renameFolders: false,
          });
        }),
      );
  }

  onClose(): void {
    const { contentEl } = this;
    contentEl.empty();
    if (this.resolver) {
      this.resolver({
        applyTypeUpdate: false,
        renameFolders: false,
      });
      this.resolver = null;
    }
  }

  private finish(result: CategoryRenameDecision): void {
    const resolve = this.resolver;
    this.resolver = null;
    this.close();
    if (resolve) {
      resolve(result);
    }
  }
}

class CreateCampaignFoldersModal extends Modal {
  private readonly missingCount: number;
  private readonly campaignName: string;
  private resolver: ((value: boolean) => void) | null = null;

  constructor(app: App, missingCount: number, campaignName: string) {
    super(app);
    this.missingCount = missingCount;
    this.campaignName = campaignName;
    this.setTitle("Create Category Folders");
  }

  static async prompt(app: App, missingCount: number, campaignName: string): Promise<boolean> {
    const modal = new CreateCampaignFoldersModal(app, missingCount, campaignName);
    return modal.openAndWait();
  }

  private openAndWait(): Promise<boolean> {
    return new Promise((resolve) => {
      this.resolver = resolve;
      this.open();
    });
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl("p", {
      text: `${this.missingCount} category folder(s) are missing for campaign "${this.campaignName}".`,
    });
    contentEl.createEl("p", {
      text: "Create them now? Existing folders will be left untouched.",
    });

    new Setting(contentEl)
      .addButton((btn) =>
        btn.setButtonText("Create folders").setCta().onClick(() => this.finish(true)),
      )
      .addButton((btn) => btn.setButtonText("Skip").onClick(() => this.finish(false)));
  }

  onClose(): void {
    const { contentEl } = this;
    contentEl.empty();
    if (this.resolver) {
      this.resolver(false);
      this.resolver = null;
    }
  }

  private finish(result: boolean): void {
    const resolve = this.resolver;
    this.resolver = null;
    this.close();
    if (resolve) {
      resolve(result);
    }
  }
}

class AdoptExistingEntityModal extends Modal {
  private readonly name: string;
  private readonly path: string;
  private resolver: ((value: boolean) => void) | null = null;

  constructor(app: App, data: { name: string; path: string }) {
    super(app);
    this.name = data.name;
    this.path = data.path;
    this.setTitle("Adopt Existing File");
  }

  static async prompt(app: App, data: { name: string; path: string }): Promise<boolean> {
    const modal = new AdoptExistingEntityModal(app, data);
    return modal.openAndWait();
  }

  private openAndWait(): Promise<boolean> {
    return new Promise((resolve) => {
      this.resolver = resolve;
      this.open();
    });
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();
    contentEl.createEl("p", {
      text: `A note named "${this.name}" already exists in this campaign.`,
    });
    contentEl.createEl("p", {
      text: this.path,
    });
    contentEl.createEl("p", {
      text: "Adopt it by adding missing frontmatter keys only, or create a new suffixed note?",
    });

    new Setting(contentEl)
      .addButton((btn) => btn.setButtonText("Adopt existing").setCta().onClick(() => this.finish(true)))
      .addButton((btn) => btn.setButtonText("Create suffixed").onClick(() => this.finish(false)));
  }

  onClose(): void {
    const { contentEl } = this;
    contentEl.empty();
    if (this.resolver) {
      this.resolver(false);
      this.resolver = null;
    }
  }

  private finish(result: boolean): void {
    const resolve = this.resolver;
    this.resolver = null;
    this.close();
    if (resolve) {
      resolve(result);
    }
  }
}

class CategorizeDocumentModal extends Modal {
  private readonly documentName: string;
  private readonly entityTypes: string[];
  private selectedType: string;
  private aliasInput = "";
  private resolver: ((value: CategorizeDocumentDecision | null) => void) | null = null;

  constructor(app: App, documentName: string, entityTypes: string[]) {
    super(app);
    this.documentName = documentName;
    this.entityTypes = entityTypes;
    this.selectedType = entityTypes[0] ?? "";
    this.setTitle("Categorize Active Document");
  }

  static async prompt(
    app: App,
    documentName: string,
    entityTypes: string[],
  ): Promise<CategorizeDocumentDecision | null> {
    const modal = new CategorizeDocumentModal(app, documentName, entityTypes);
    return modal.openAndWait();
  }

  private openAndWait(): Promise<CategorizeDocumentDecision | null> {
    return new Promise((resolve) => {
      this.resolver = resolve;
      this.open();
    });
  }

  onOpen(): void {
    const { contentEl } = this;
    contentEl.empty();

    contentEl.createEl("p", {
      text: `Categorize "${this.documentName}" and prepare alias-based link backfill.`,
    });

    new Setting(contentEl)
      .setName("Category")
      .setDesc("Choose the entity category for this document.")
      .addDropdown((dropdown) => {
        this.entityTypes.forEach((entityType) => dropdown.addOption(entityType, entityType));
        dropdown.setValue(this.selectedType).onChange((value) => {
          this.selectedType = normalizeEntityType(value);
        });
      });

    new Setting(contentEl)
      .setName("Alias seeds")
      .setDesc("Optional comma, semicolon, or newline-separated aliases to add before scanning campaign prose.")
      .addTextArea((text) => {
        text.setPlaceholder("Ari, The Iron Warden").setValue(this.aliasInput).onChange((value) => {
          this.aliasInput = value;
        });
        text.inputEl.rows = 4;
        text.inputEl.cols = 32;
      });

    new Setting(contentEl)
      .addButton((btn) =>
        btn.setButtonText("Categorize").setCta().onClick(() => {
          if (!this.selectedType) {
            new Notice("Select a category.");
            return;
          }
          this.finish({
            entityType: normalizeEntityType(this.selectedType),
            aliases: parseCustomSpellings(this.aliasInput),
          });
        }),
      )
      .addButton((btn) => btn.setButtonText("Cancel").onClick(() => this.finish(null)));
  }

  onClose(): void {
    const { contentEl } = this;
    contentEl.empty();
    if (this.resolver) {
      this.resolver(null);
      this.resolver = null;
    }
  }

  private finish(result: CategorizeDocumentDecision | null): void {
    const resolve = this.resolver;
    this.resolver = null;
    this.close();
    if (resolve) {
      resolve(result);
    }
  }
}

class VariantBackfillModal extends Modal {
  private readonly canonicalName: string;
  private matches: VariantMatch[];
  private readonly expandMatches?: VariantMatchExpander;
  private resolver: ((value: VariantBackfillDecision) => void) | null = null;

  private applyLinks = true;
  private addAliases = true;
  private selectedVariants: Set<string>;
  private customSpellingInput = "";

  constructor(app: App, canonicalName: string, matches: VariantMatch[], expandMatches?: VariantMatchExpander) {
    super(app);
    this.canonicalName = canonicalName;
    this.matches = matches;
    this.expandMatches = expandMatches;
    this.selectedVariants = new Set(matches.map((m) => m.variant));
    this.setTitle("Link Existing Name Variants");
  }

  static async prompt(
    app: App,
    canonicalName: string,
    matches: VariantMatch[],
    expandMatches?: VariantMatchExpander,
  ): Promise<VariantBackfillDecision> {
    const modal = new VariantBackfillModal(app, canonicalName, matches, expandMatches);
    return modal.openAndWait();
  }

  private openAndWait(): Promise<VariantBackfillDecision> {
    return new Promise((resolve) => {
      this.resolver = resolve;
      this.open();
    });
  }

  onOpen(): void {
    this.render();
  }

  private render(): void {
    const { contentEl } = this;
    contentEl.empty();

    const totalHits = this.matches.reduce((sum, m) => sum + m.count, 0);
    contentEl.createEl("p", {
      text: `Found ${totalHits} possible occurrence(s) of "${this.canonicalName}" across ${this.matches.length} spelling variant(s).`,
    });

    const preview = contentEl.createEl("ul");
    this.matches.slice(0, 8).forEach((m) => {
      preview.createEl("li", { text: `${m.variant} (${m.count})` });
    });
    if (this.matches.length > 8) {
      contentEl.createEl("small", { text: `+ ${this.matches.length - 8} more variant(s)` });
    }

    new Setting(contentEl)
      .setName("Custom spelling seeds")
      .setDesc("Comma-separated extra forms to widen fuzzy search (for example: Ari, Arri).")
      .addText((text) =>
        text.setPlaceholder("Ari, Arri").setValue(this.customSpellingInput).onChange((value) => {
          this.customSpellingInput = value;
        }),
      )
      .addButton((btn) =>
        btn.setButtonText("Expand search").onClick(async () => {
          if (!this.expandMatches) return;
          const customSpellings = parseCustomSpellings(this.customSpellingInput);
          if (customSpellings.length === 0) {
            new Notice("Add at least one custom spelling.");
            return;
          }
          btn.setDisabled(true);
          btn.setButtonText("Searching...");
          try {
            const expanded = await this.expandMatches(customSpellings);
            this.matches = expanded;
            for (const match of expanded) {
              this.selectedVariants.add(match.variant);
            }
            this.render();
          } finally {
            btn.setDisabled(false);
            btn.setButtonText("Expand search");
          }
        }),
      );

    new Setting(contentEl)
      .setName("Replace variant occurrences with links")
      .setDesc("Preserve prose text via aliased links.")
      .addToggle((toggle) =>
        toggle.setValue(this.applyLinks).onChange((value) => {
          this.applyLinks = value;
        }),
      );

    new Setting(contentEl)
      .setName("Add divergent variants as aliases")
      .setDesc("Only variants that differ from the canonical name are added.")
      .addToggle((toggle) =>
        toggle.setValue(this.addAliases).onChange((value) => {
          this.addAliases = value;
        }),
      );

    contentEl.createEl("h4", { text: "Variant selection" });
    contentEl.createEl("small", {
      text: "Only checked variants are processed for links/aliases.",
    });

    this.matches.forEach((m) => {
      new Setting(contentEl)
        .setName(`${m.variant}`)
        .setDesc(`${m.count} occurrence(s)`)
        .addToggle((toggle) =>
          toggle.setValue(this.selectedVariants.has(m.variant)).onChange((value) => {
            if (value) {
              this.selectedVariants.add(m.variant);
            } else {
              this.selectedVariants.delete(m.variant);
            }
          }),
        );
    });

    new Setting(contentEl)
      .addButton((btn) =>
        btn.setButtonText("Apply").setCta().onClick(() => {
          this.finish({
            applyLinks: this.applyLinks,
            addAliases: this.addAliases,
            selectedVariants: Array.from(this.selectedVariants),
          });
        }),
      )
      .addButton((btn) =>
        btn.setButtonText("Skip").onClick(() => {
          this.finish({
            applyLinks: false,
            addAliases: false,
            selectedVariants: [],
          });
        }),
      );
  }

  onClose(): void {
    const { contentEl } = this;
    contentEl.empty();
    if (this.resolver) {
      this.resolver({
        applyLinks: false,
        addAliases: false,
        selectedVariants: [],
      });
      this.resolver = null;
    }
  }

  private finish(result: VariantBackfillDecision): void {
    const resolve = this.resolver;
    this.resolver = null;
    this.close();
    if (resolve) {
      resolve(result);
    }
  }
}

function parseMentionAtEnd(
  linePrefix: string,
  configuredPrefixes: string[],
  modes: MentionInputModes,
): { prefix: string; rawQuery: string; startCh: number } | null {
  if (configuredPrefixes.length === 0) return null;
  const sortedPrefixes = [...configuredPrefixes].sort((a, b) => b.length - a.length);

  for (let i = linePrefix.length - 1; i >= 0; i -= 1) {
    const boundaryOkay = i === 0 || /\s/.test(linePrefix[i - 1] ?? "");
    if (!boundaryOkay) continue;

    const prefix = sortedPrefixes.find((candidate) => linePrefix.startsWith(candidate, i));
    if (!prefix) continue;

    const startCh = i;
    const afterPrefix = linePrefix.slice(i + prefix.length);
    const trimmedStart = afterPrefix.replace(/^\s+/, "");
    if (!trimmedStart) return null;

    if (modes.quoted && trimmedStart.startsWith(`"`)) {
      const inner = trimmedStart.slice(1);
      const closingQuote = inner.lastIndexOf(`"`);
      if (closingQuote >= 0 && closingQuote === inner.length - 1) {
        const rawQuery = inner.slice(0, -1).trim();
        if (rawQuery) return { prefix, rawQuery, startCh };
      } else {
        const rawQuery = inner.trim();
        if (rawQuery) return { prefix, rawQuery, startCh };
      }
    }

    if (modes.open) {
      const openText = trimmedStart.trimEnd();
      if (!openText) return null;
      if (/[.,;:!?()[\]{}]$/.test(openText)) return null;
      return { prefix, rawQuery: openText, startCh };
    }

    if (modes.singleWord) {
      const singleWordMatch = trimmedStart.match(/^([^\s\[\]{}()<>.,;:!?]+)/);
      const rawQuery = singleWordMatch?.[1]?.trim() ?? "";
      if (rawQuery) return { prefix, rawQuery, startCh };
    }
  }

  return null;
}

function parseMentionQuery(query: string): ParsedMentionQuery | null {
  try {
    const parsed = JSON.parse(query) as ParsedMentionQuery;
    if (!parsed.prefix || !parsed.entityType) return null;
    return parsed;
  } catch {
    return null;
  }
}

function parseTrailingWordFragment(linePrefix: string): { text: string; startCh: number } | null {
  const match = linePrefix.match(/([A-Za-z0-9][A-Za-z0-9'-]*)$/);
  if (!match?.[1]) return null;
  const text = match[1];
  const startCh = linePrefix.length - text.length;
  if (startCh > 0) {
    const prev = linePrefix[startCh - 1] ?? "";
    if (/[A-Za-z0-9\]`]/.test(prev)) return null;
  }
  return { text, startCh };
}

function isImmediatelyAfterConfiguredPrefix(beforeFragment: string, prefixes: string[]): boolean {
  const trimmed = beforeFragment.replace(/\s+$/, "");
  if (!trimmed) return false;
  return prefixes.some((prefix) => trimmed.endsWith(prefix));
}

function clampNumber(value: unknown, min: number, max: number, fallback: number): number {
  const num = typeof value === "number" ? value : Number.parseInt(String(value), 10);
  if (!Number.isFinite(num)) return fallback;
  if (num < min) return min;
  if (num > max) return max;
  return Math.floor(num);
}

function scoreProseSuggestion(suggestion: ProseMentionSuggestion, query: string): number {
  const candidate = suggestion.normalizedDisplay;
  if (candidate === query) return suggestion.isAlias ? 1 : 0;
  if (candidate.startsWith(query)) return suggestion.isAlias ? 3 : 2;
  if (candidate.includes(query)) return 10;
  return 100;
}

function normalizeSearchText(value: string): string {
  return value.toLowerCase().replace(/\s+/g, " ").trim();
}

function normalizeEntityType(value: string): string {
  return value.trim().toLowerCase();
}

function normalizeHexColor(value: unknown): string | undefined {
  if (typeof value !== "string") return undefined;
  const trimmed = value.trim();
  if (!/^#[0-9a-fA-F]{6}$/.test(trimmed)) return undefined;
  return trimmed.toLowerCase();
}

function scoreRecord(record: EntityRecord, query: string): number {
  if (!query) return 1000;
  if (record.normalizedName === query) return 0;
  if (record.normalizedAliases.includes(query)) return 1;
  if (record.normalizedName.startsWith(query)) return 2;
  if (record.normalizedAliases.some((alias) => alias.startsWith(query))) return 3;
  return 10;
}

function sanitizeFileName(value: string): string {
  return value.replace(/[\\/:*?"<>|#^[\]]/g, " ").replace(/\s+/g, " ").trim();
}

function normalizeWhitespace(value: string): string {
  return value.replace(/\s+/g, " ").trim();
}

function parseCustomSpellings(value: string): string[] {
  if (!value.trim()) return [];
  return Array.from(
    new Set(
      value
        .split(/[,\n;]+/)
        .map((part) => normalizeWhitespace(part))
        .filter((part) => part.length > 0),
    ),
  );
}

function normalizeFuzzyText(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9\s'-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function splitWords(value: string): string[] {
  return value.split(/\s+/).filter(Boolean);
}

function* iterateProseLines(content: string): Generator<string> {
  const lines = content.split(/\r?\n/);
  let inFrontmatter = false;
  let frontmatterHandled = false;
  let inFence = false;

  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i] ?? "";

    if (!frontmatterHandled && i === 0 && line.trim() === "---") {
      inFrontmatter = true;
      continue;
    }
    if (inFrontmatter) {
      if (line.trim() === "---") {
        inFrontmatter = false;
        frontmatterHandled = true;
      }
      continue;
    }
    frontmatterHandled = true;

    if (/^\s*```/.test(line)) {
      inFence = !inFence;
      continue;
    }
    if (inFence) continue;

    yield line;
  }
}

function collectFuzzyCandidatesFromLine(line: string, targetWordCount: number): string[] {
  const matches = Array.from(line.matchAll(/\b[a-zA-Z0-9][a-zA-Z0-9'’-]*\b/g));
  if (matches.length === 0) return [];

  if (targetWordCount <= 1) {
    return Array.from(
      new Set(
        matches
          .map((m) => (m[0] ?? "").trim())
          .filter((token) => token.length >= 2),
      ),
    );
  }

  const windows = new Set<string>();
  const minLen = Math.max(1, targetWordCount - 1);
  const maxLen = targetWordCount + 2;

  for (let i = 0; i < matches.length; i += 1) {
    for (let len = minLen; len <= maxLen; len += 1) {
      const endIdx = i + len - 1;
      if (endIdx >= matches.length) continue;
      const start = matches[i]?.index;
      const endMatch = matches[endIdx];
      if (start === undefined || endMatch?.index === undefined) continue;
      const end = endMatch.index + endMatch[0].length;
      const phrase = line.slice(start, end).trim();
      if (!phrase || phrase.length < 3) continue;
      windows.add(phrase);
    }
  }

  return Array.from(windows);
}

function looksLikeSameEntity(target: string, candidate: string): boolean {
  if (!target || !candidate) return false;
  if (target === candidate) return true;
  const targetWords = splitWords(target);
  const candidateWords = splitWords(candidate);

  if (targetWords.length === 1) {
    if (candidateWords.length !== 1) return false;
    const t = targetWords[0] ?? "";
    const c = candidateWords[0] ?? "";
    if (!t || !c) return false;
    const minLen = Math.min(t.length, c.length);
    if (minLen <= 3) return t === c;
    if (t[0] !== c[0]) return false;
    if (Math.abs(t.length - c.length) > 2) return false;
    const distance = levenshtein(t, c);
    if (distance > 1) return false;
    if (minLen <= 4) {
      return t[t.length - 1] === c[c.length - 1];
    }
    const prefixLen = commonPrefixLength(t, c);
    if (prefixLen >= Math.max(2, minLen - 2)) return true;
    return t[t.length - 1] === c[c.length - 1];
  }

  if (target[0] !== candidate[0]) return false;
  const overlap = candidateWords.filter((w) => targetWords.includes(w)).length;
  return overlap >= Math.max(1, Math.min(targetWords.length, candidateWords.length) - 1);
}

function maxLevenshteinFor(candidate: string): number {
  if (candidate.length <= 8) return 1;
  if (candidate.length <= 16) return 2;
  return 3;
}

function levenshtein(a: string, b: string): number {
  if (a === b) return 0;
  if (!a) return b.length;
  if (!b) return a.length;

  const prev = new Array<number>(b.length + 1);
  const curr = new Array<number>(b.length + 1);

  for (let j = 0; j <= b.length; j += 1) prev[j] = j;

  for (let i = 1; i <= a.length; i += 1) {
    curr[0] = i;
    for (let j = 1; j <= b.length; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min((curr[j - 1] ?? 0) + 1, (prev[j] ?? 0) + 1, (prev[j - 1] ?? 0) + cost);
    }
    for (let j = 0; j <= b.length; j += 1) prev[j] = curr[j] ?? 0;
  }

  return prev[b.length] ?? 0;
}

function commonPrefixLength(a: string, b: string): number {
  const max = Math.min(a.length, b.length);
  let i = 0;
  while (i < max && a[i] === b[i]) i += 1;
  return i;
}

function replaceVariantOccurrencesInText(
  content: string,
  variants: string[],
  linkPath: string,
  onReplace: () => void,
): string {
  const lines = content.split(/\r?\n/);
  const endings = content.includes("\r\n") ? "\r\n" : "\n";
  let inFrontmatter = false;
  let frontmatterHandled = false;
  let inFence = false;

  const replacedLines = lines.map((line, lineIndex) => {
    if (!frontmatterHandled && lineIndex === 0 && line.trim() === "---") {
      inFrontmatter = true;
      return line;
    }
    if (inFrontmatter) {
      if (line.trim() === "---") {
        inFrontmatter = false;
        frontmatterHandled = true;
      }
      return line;
    }
    frontmatterHandled = true;

    if (/^\s*```/.test(line)) {
      inFence = !inFence;
      return line;
    }
    if (inFence) return line;

    const protectedPattern = /(\[\[[^\]\n]+\]\]|\[[^\]\n]+\]\([^)]+\)|`[^`\n]*`)/g;
    const segments = line.split(protectedPattern);
    const updatedSegments = segments.map((segment) => {
      if (isProtectedSegment(segment)) return segment;

      let next = segment;
      for (const variant of variants) {
        const parts = next.split(protectedPattern);
        const escaped = escapeRegex(variant);
        const pattern = new RegExp(`(^|[^A-Za-z0-9])(${escaped})(?=$|[^A-Za-z0-9])`, "g");
        next = parts
          .map((part) => {
            if (isProtectedSegment(part)) return part;
            return part.replace(pattern, (_full, leading, matched) => {
              onReplace();
              return `${leading}[[${linkPath}|${matched}]]`;
            });
          })
          .join("");
      }
      return next;
    });

    return updatedSegments.join("");
  });

  return replacedLines.join(endings);
}

function isProtectedSegment(value: string): boolean {
  if (!value) return false;
  if (/^\[\[[^\]\n]+\]\]$/.test(value)) return true;
  if (/^\[[^\]\n]+\]\([^)]+\)$/.test(value)) return true;
  if (/^`[^`\n]*`$/.test(value)) return true;
  return false;
}

function normalizeHotkeySetting(value: string): string {
  const trimmed = value.trim().toLowerCase();
  if (!trimmed) return "";
  const parts = trimmed
    .split("+")
    .map((part) => part.trim())
    .filter(Boolean);
  if (parts.length === 0) return "";

  let useMod = false;
  let ctrl = false;
  let meta = false;
  let alt = false;
  let shift = false;
  let key = "";

  for (const part of parts) {
    if (part === "mod") {
      useMod = true;
      continue;
    }
    if (part === "ctrl" || part === "control") {
      ctrl = true;
      continue;
    }
    if (part === "meta" || part === "cmd" || part === "command" || part === "win" || part === "windows" || part === "super") {
      meta = true;
      continue;
    }
    if (part === "alt" || part === "option") {
      alt = true;
      continue;
    }
    if (part === "shift") {
      shift = true;
      continue;
    }
    if (!key) {
      key = part;
    }
  }

  const out: string[] = [];
  if (useMod) out.push("mod");
  if (ctrl) out.push("ctrl");
  if (meta) out.push("meta");
  if (alt) out.push("alt");
  if (shift) out.push("shift");
  if (key) out.push(key);
  return out.join("+");
}

function formatHotkeyFromEvent(evt: KeyboardEvent): string | null {
  const key = evt.key.toLowerCase();
  if (key === "control" || key === "shift" || key === "alt" || key === "meta") return null;
  if (!evt.ctrlKey && !evt.metaKey && !evt.altKey && !evt.shiftKey) return null;

  const parts: string[] = [];
  if (evt.ctrlKey) parts.push("ctrl");
  if (evt.metaKey) parts.push("meta");
  if (evt.altKey) parts.push("alt");
  if (evt.shiftKey) parts.push("shift");

  if (key === " ") parts.push("space");
  else parts.push(key);
  return parts.join("+");
}

function hotkeyMatchesEvent(hotkey: string, evt: KeyboardEvent): boolean {
  const normalized = normalizeHotkeySetting(hotkey);
  if (!normalized) return false;
  const parts = normalized.split("+");
  if (parts.length === 0) return false;

  const key = parts[parts.length - 1] ?? "";
  const modifiers = new Set(parts.slice(0, -1));
  const isMac = typeof navigator !== "undefined" && /mac/i.test(navigator.platform);

  const expectCtrl = modifiers.has("ctrl") || (modifiers.has("mod") && !isMac);
  const expectMeta = modifiers.has("meta") || (modifiers.has("mod") && isMac);
  const expectAlt = modifiers.has("alt");
  const expectShift = modifiers.has("shift");

  if (evt.ctrlKey !== expectCtrl) return false;
  if (evt.metaKey !== expectMeta) return false;
  if (evt.altKey !== expectAlt) return false;
  if (evt.shiftKey !== expectShift) return false;

  const eventKey = evt.key.toLowerCase();
  if (key === "space") return eventKey === " ";
  return eventKey === key;
}

function normalizeLinkPathForResolve(rawLink: string): string | null {
  let value = rawLink.trim();
  if (!value) return null;

  if (value.startsWith("[[") && value.endsWith("]]")) {
    value = value.slice(2, -2).trim();
  }

  const pipeIndex = value.indexOf("|");
  if (pipeIndex >= 0) value = value.slice(0, pipeIndex);

  const hashIndex = value.indexOf("#");
  if (hashIndex >= 0) value = value.slice(0, hashIndex);

  const blockRefIndex = value.indexOf("^");
  if (blockRefIndex >= 0) value = value.slice(0, blockRefIndex);

  try {
    value = decodeURIComponent(value);
  } catch {
    // Ignore malformed URI fragments and continue with raw value.
  }

  value = value.trim();
  return value || null;
}

function slugify(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "") || "campaign";
}

function toCssSlug(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function isPathInScope(path: string, scopeRoot: string): boolean {
  const normalizedPath = normalizePath(path);
  const normalizedRoot = normalizePath(scopeRoot).replace(/\/+$/g, "");
  return normalizedPath === normalizedRoot || normalizedPath.startsWith(`${normalizedRoot}/`);
}

function frontmatterAliasesToArray(raw: unknown): string[] {
  if (Array.isArray(raw)) {
    return raw.map((v) => String(v)).filter(Boolean);
  }
  if (typeof raw === "string") {
    return raw
      .split(",")
      .map((v) => v.trim())
      .filter(Boolean);
  }
  return [];
}

function defaultFolderForType(entityType: string, capitalize: boolean): string {
  const type = normalizeEntityType(entityType);
  return formatCategoryFolderName(type, capitalize);
}

function formatCategoryFolderName(value: string, capitalize: boolean): string {
  const clean = value.trim();
  if (!clean) return "";
  if (!capitalize) return clean;

  if (clean.toLowerCase() === "npcs") return "NPCs";
  if (clean.toLowerCase() === "pcs") return "PCs";

  return clean
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1))
    .join(" ");
}

async function ensureFolder(app: App, folderPath: string): Promise<void> {
  const normalized = normalizePath(folderPath);
  if (app.vault.getAbstractFileByPath(normalized)) return;

  const parts = normalized.split("/");
  let current = "";
  for (const part of parts) {
    current = current ? `${current}/${part}` : part;
    if (!app.vault.getAbstractFileByPath(current)) {
      await app.vault.createFolder(current);
    }
  }
}
