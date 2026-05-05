namespace PowerBI.Models
{
    public class Workspace
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? PowerBIWorkspaceId { get; set; }
        public int UserId { get; set; }
    }
}
