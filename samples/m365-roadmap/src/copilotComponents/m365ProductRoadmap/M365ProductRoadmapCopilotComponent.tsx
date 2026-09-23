import * as React from "react";
import { createRoot, type Root } from "react-dom/client";
import { BaseCopilotComponent } from "@microsoft/sp-copilot-component";
import type { Client } from "@modelcontextprotocol/sdk/client/index.js";

import M365ProductRoadmapMcpApp from "./components/M365ProductRoadmapMcpApp";
import type { IM365ProductRoadmapMcpAppProps } from "./components/IM365ProductRoadmapMcpAppProps";
import type { IM365ProductRoadmapCopilotComponentProperties } from "./M365ProductRoadmapCopilotComponentProperties";

export default class M365ProductRoadmapCopilotComponent extends BaseCopilotComponent<IM365ProductRoadmapCopilotComponentProperties> {
  private _root: Root | undefined;
  private _activeClient: Client | undefined;
  private _toolInput: Record<string, unknown> = {};

  private readonly _handleClientChanged = (
    client: Client | undefined,
  ): void => {
    this._activeClient = client;
  };

  protected async onInit(): Promise<void> {
    const properties = this.properties;
    this._toolInput = {
      ...(properties.searchQuery
        ? { searchQuery: properties.searchQuery }
        : {}),
      ...(properties.roadmapId ? { roadmapId: properties.roadmapId } : {}),
      ...(properties.products?.length ? { products: properties.products } : {}),
      ...(properties.statuses?.length ? { statuses: properties.statuses } : {}),
      ...(properties.platforms?.length
        ? { platforms: properties.platforms }
        : {}),
      ...(properties.releaseRings?.length
        ? { releaseRings: properties.releaseRings }
        : {}),
      ...(properties.cloudInstances?.length
        ? { cloudInstances: properties.cloudInstances }
        : {}),
      ...(properties.availabilityFrom
        ? { availabilityFrom: properties.availabilityFrom }
        : {}),
      ...(properties.availabilityTo
        ? { availabilityTo: properties.availabilityTo }
        : {}),
      ...(properties.modifiedFrom
        ? { modifiedFrom: properties.modifiedFrom }
        : {}),
      ...(properties.modifiedTo ? { modifiedTo: properties.modifiedTo } : {}),
    };
    // Request fullscreen on initialization
    try {
      await this.requestDisplayModeAsync("fullscreen");
      console.log("[M365ProductRoadmap] Requested fullscreen mode on init");
    } catch (error) {
      console.warn(
        "[M365ProductRoadmap] Failed to request fullscreen on init:",
        error,
      );
      // The component will still render in inline mode  
    }
  }

  protected render(): void {
    const props: IM365ProductRoadmapMcpAppProps = {
      assetBaseUrl:
        this.context.manifest.loaderConfig.internalModuleBaseUrls[0],
      bridge: this.context.copilotBridge,
      hostContext: this.hostContext,
      locale:
        this.context.pageContext.cultureInfo.currentUICultureName || "en-US",
      onClientChanged: this._handleClientChanged,
      onRequestDisplayMode: async (mode) => this.requestDisplayModeAsync(mode),

      targetDocument: this.context.domElement.ownerDocument,
      toolInput: this._toolInput,
    };

    if (!this._root) {
      this._root = createRoot(this.context.domElement);
    }

    this._root.render(React.createElement(M365ProductRoadmapMcpApp, props));
  }

  protected async onTeardown(_reason: string | undefined): Promise<void> {
    const activeClient: Client | undefined = this._activeClient;
    this._root?.unmount();
    this._root = undefined;
    this._activeClient = undefined;
    await activeClient?.close().catch(() => undefined);
  }
}
