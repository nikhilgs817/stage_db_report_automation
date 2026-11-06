// Paste in the SCRIPT_SELECTION Active Choices Parameter in the job UI (Single Select).
// No referenced parameters. Untick "Use Groovy Sandbox"; the Jenkins admin approves it.
// Lists committed .sql files on REF from the GitLab API, pinned to REF's current commit.
// It never connects to DB2. Set the four values below for your environment.
import groovy.json.JsonSlurper
import jenkins.model.Jenkins
import com.cloudbees.plugins.credentials.CredentialsProvider
import org.jenkinsci.plugins.plaincredentials.StringCredentials

def GITLAB_URL = 'https://gitlab.example.com'
def PROJECT = 'group%2Freport-sql'   // URL-encoded project path, or the numeric project ID
def REF = 'main'                     // approved branch of the SQL repository
def TOKEN_ID = 'gitlab-api-token'    // Secret text credential (read_api), stored at global scope

try {
    def credential = CredentialsProvider.lookupCredentials(StringCredentials, Jenkins.get(), null, null)
        .find { it.id == TOKEN_ID }
    if (credential == null) return ["Credential ${TOKEN_ID} not found; ask Jenkins administrator:disabled".toString()]
    def api = "${GITLAB_URL}/api/v4/projects/${PROJECT}/repository"
    def get = { String url ->
        def connection = new URL(url).openConnection()
        connection.setRequestProperty('PRIVATE-TOKEN', credential.secret.plainText)
        connection.connectTimeout = 5000
        connection.readTimeout = 15000
        if (connection.responseCode != 200) throw new IOException("GitLab HTTP ${connection.responseCode}")
        [new JsonSlurper().parse(connection.inputStream, 'UTF-8'), connection.getHeaderField('X-Next-Page')]
    }
    def commit = get("${api}/commits/${REF}")[0].id?.toString()
    if (!(commit ==~ /[0-9a-f]{40}/)) return ['GitLab returned an invalid commit; ask Jenkins administrator:disabled']
    def files = []
    def page = '1'
    while (page) {
        def (rows, next) = get("${api}/tree?ref=${commit}&recursive=true&per_page=100&page=${page}")
        files += rows.findAll { it.type == 'blob' && it.path?.toLowerCase()?.endsWith('.sql') }*.path
        page = next
    }
    def choices = files.findAll { path ->
        !path.contains('|') && !path.contains('\n') && !path.contains('\r')
    }.sort().collect { "${it}|${commit}".toString() }
    return choices ?: ['No SQL scripts found:disabled']
} catch (IOException failure) {
    return ["GitLab unavailable (${failure.message}); ask Jenkins administrator:disabled".toString()]
} catch (Exception ignored) {
    return ['SQL list unavailable; ask Jenkins administrator:disabled']
}
