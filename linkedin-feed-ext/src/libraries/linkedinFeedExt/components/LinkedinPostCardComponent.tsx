import * as React from "react";
import * as ReactDOM from "react-dom";
import { BaseWebComponent } from "@pnp/modern-search-extensibility";
import { LinkedinPostCard } from "./LinkedinPostCard";

export class LinkedinPostCardComponent extends BaseWebComponent {
  public async connectedCallback(): Promise<void> {
    const props = this.resolveAttributes();

    // We expect the proxy API to return the post content and an author.
    // Ensure all props are strings when passed via Handlebars HTML attributes.
    ReactDOM.render(
      React.createElement(LinkedinPostCard, {
        authorName: props.authorName as string || "LinkedIn User",
        authorTitle: props.authorTitle as string || "",
        postText: props.postText as string || "",
        postUrl: props.postUrl as string || "#",
        imageUrl: props.imageUrl as string || "",
        date: props.date as string || ""
      }),
      this
    );
  }

  protected disconnectedCallback(): void {
    ReactDOM.unmountComponentAtNode(this);
  }
}
