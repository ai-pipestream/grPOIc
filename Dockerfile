# Build stage: compile, run the full test suite, and assemble the
# distribution. An image never ships from a tree whose tests did not pass.
# The -jdk-dev variant of the Docker Hardened Images temurin carries a
# shell and runs as root, which the Gradle build needs.
FROM dhi.io/eclipse-temurin:25-jdk-dev AS build
WORKDIR /src
COPY . .
RUN ./gradlew --no-daemon build :grpoic-service:installDist

# Runtime: Docker Hardened Images temurin JRE. No shell, no package
# manager, runs as non-root uid 65532 out of the box (no useradd/USER
# needed). The server is diskless by doctrine; run with --read-only and it
# works unchanged. With no shell the Gradle start script cannot run, so
# java is invoked directly on the installDist classpath.
FROM dhi.io/eclipse-temurin:25
COPY --from=build /src/grpoic-service/build/install/grpoic-service /opt/grpoic
EXPOSE 50052
ENTRYPOINT ["java", "-cp", "/opt/grpoic/lib/*", "ai.pipestream.grpoic.server.GrPoicServer"]
