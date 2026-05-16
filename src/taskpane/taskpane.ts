/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, process */

import * as wordService from "../office/wordService";
import { initApp } from "../ui/app";
import { enableSoftHmr } from "./devReload";

if (process.env.NODE_ENV === "development") {
  enableSoftHmr();
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Word) {
    return;
  }

  initApp(wordService);
});
