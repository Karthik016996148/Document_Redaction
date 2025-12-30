import { defineConfig } from "vite";
import { getHttpsServerOptions } from "office-addin-dev-certs";

export default defineConfig(async () => {
  let https: boolean | Record<string, unknown> = true;
  try {
    https = await getHttpsServerOptions();
  } catch {
    // If dev-certs aren't installed yet, still start; Word may prompt/trust later.
    https = true;
  }

  return {
    server: {
      port: 3000,
      strictPort: true,
      https,
      cors: true,
    },
    preview: {
      port: 3000,
      strictPort: true,
      https,
    },
  };
});


