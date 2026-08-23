plugins {
    application
}

dependencies {
    implementation(project(":grpoic-api"))
    implementation(libs.grpc.netty.shaded)
    implementation(libs.grpc.services)
    implementation(libs.poi)
    implementation(libs.poi.ooxml)
    implementation(libs.poi.scratchpad)
    runtimeOnly(libs.log4j.core)

    testImplementation(libs.grpc.inprocess)
    testImplementation(libs.junit.jupiter)
    testImplementation(libs.assertj.core)
    testImplementation(libs.awaitility)
    testRuntimeOnly(libs.junit.launcher)
}

application {
    mainClass = "ai.pipestream.grpoic.server.GrPoicServer"
}
