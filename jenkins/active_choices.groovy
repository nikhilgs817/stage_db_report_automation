// Paste in the SCRIPT_SELECTION Active Choices Reactive Parameter in the JOB UI.
// Referenced parameters: CATALOG_JOB. Review/approve this script with Jenkins admin.
// It reads one published metadata artifact; it never connects to Git or DB2.
import groovy.json.JsonSlurper
import hudson.Functions
import hudson.model.Item
import hudson.model.Run
import jenkins.model.Jenkins

try {
    def jobName = binding.hasVariable('CATALOG_JOB') ? CATALOG_JOB?.toString()?.trim() : ''
    if (!jobName) return ['Configure CATALOG_JOB first:disabled']
    def job = Jenkins.get().getItemByFullName(jobName)
    if (job == null) return ['Catalog job unavailable:disabled']
    job.checkPermission(Item.READ)
    def run = job.getLastSuccessfulBuild()
    if (run == null) return ['Run the catalog job first:disabled']
    if (Functions.isArtifactsPermissionEnabled()) run.checkPermission(Run.ARTIFACTS)
    def artifact = run.getArtifactManager().root().child('sql-catalog.json')
    if (!artifact.exists()) return ['Refresh the SQL catalog:disabled']
    def catalog
    artifact.open().withCloseable { stream ->
        catalog = new JsonSlurper().parse(stream, 'UTF-8')
    }
    if (catalog.version != 1 || !(catalog.commit ==~ /[0-9a-f]{40}/) ||
            !(catalog.choices instanceof List)) {
        return ['Invalid catalog; refresh required:disabled']
    }
    def choices = catalog.choices.findAll { value ->
        value instanceof String && value.endsWith('|' + catalog.commit) &&
            !value.contains('\n') && !value.contains('\r')
    }.sort()
    return choices ?: ['No SQL scripts found:disabled']
} catch (Exception ignored) {
    return ['Catalog unavailable; ask Jenkins administrator:disabled']
}
