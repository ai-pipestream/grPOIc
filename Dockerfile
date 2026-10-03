# Build stage: compile, run the full test suite, and assemble the
# distribution. An image never ships from a tree whose tests did not pass.
# Both base images are pinned by their multi-arch index digest, so a
# retagged upstream image cannot change a build; the tag stays for humans
# and for Renovate, which updates the digest.
# The -jdk-dev variant of the Docker Hardened Images temurin carries a
# shell and runs as root, which the Gradle build needs.
FROM dhi.io/eclipse-temurin:25-jdk-dev@sha256:95e456232f6a87c4291315e6a3b4fb19ba36ae7bd8f6bba6f8316e49b60dab83 AS build
WORKDIR /src
COPY . .
RUN ./gradlew --no-daemon build :grpoic-service:installDist

# Runtime: Docker Hardened Images temurin JRE. No shell, no package
# manager, runs as non-root uid 65532 out of the box (no useradd/USER
# needed). The server is diskless by doctrine; run with --read-only and it
# works unchanged. With no shell the Gradle start script cannot run, so
# java is invoked directly on the installDist classpath.
#
# Heap: 60% of the container's memory limit instead of the JVM's 25%
# default, which starves parsing. The rest covers Netty's direct buffers
# (a single-chunk upload is received off-heap before it is copied in),
# metaspace and thread stacks. No ExitOnOutOfMemoryError: a parse that
# runs out of heap is caught, its call fails with RESOURCE_EXHAUSTED and
# its document becomes garbage, while the other calls keep running.
FROM dhi.io/eclipse-temurin:25@sha256:a84cb91dd9815564076eaff9f4d762ff893c4c97a8d81cbb68d2af066798cc9c
COPY --from=build /src/grpoic-service/build/install/grpoic-service /opt/grpoic
EXPOSE 50052
ENTRYPOINT ["java", "-XX:MaxRAMPercentage=60", "-cp", "/opt/grpoic/lib/*", \
            "ai.pipestream.grpoic.server.GrPoicServer"]
