module.exports = {
    apps: [
      {
        name: 'hr-app',
        script: 'routes.py',
        interpreter: 'python',
        watch: true,
        env: {
          FLASK_ENV: 'development',
        },
        env_production: {
          FLASK_ENV: 'production',
        },
      },
    ],
  };
  