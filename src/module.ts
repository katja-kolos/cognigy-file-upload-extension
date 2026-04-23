import { createExtension } from "@cognigy/extension-tools";

import { OAuth2Connection } from "./connections/oauth2Connection";
import { tokenConnection } from "./connections/tokenConnection";
import { moveBlob } from "./nodes/flowNode";

export default createExtension({
  connections: [OAuth2Connection, tokenConnection],
  nodes: [
    moveBlob
  ],
  options: { label: "Service Now Storage" }
});