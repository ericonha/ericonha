# Use Python base image
FROM python:3.8-slim

# Set working directory
WORKDIR /app

# Copy the Python scripts into the container
COPY . /app

# Check for the operating system
RUN unameOut="$(uname -s)" \
    && case "${unameOut}" in \
        Linux*) \
            apt-get update && \
            apt-get install -y --no-install-recommends \
                build-essential && \
            rm -rf /var/lib/apt/lists/* \
            ;; \
        Darwin*) \
            echo 'Operating system is macOS' \
            ;; \
        *) \
            echo 'Unsupported operating system' \
            ;; \
    esac

# Run the Python scripts
CMD [ "python", "./main.py" ]
