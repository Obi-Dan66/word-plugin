interface NodeModule {
  hot?: {
    accept(callback?: () => void): void;
  };
}

declare const process: {
  env: {
    NODE_ENV: string;
  };
};
