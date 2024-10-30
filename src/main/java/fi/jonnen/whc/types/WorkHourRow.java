package fi.jonnen.whc.types;

import java.util.Date;
import java.util.Objects;

public class WorkHourRow implements Comparable<WorkHourRow> {
	private Date date;
	private String comments;
	private String project;
	private String task;
	private Double workHours;

	public Date getDate() {
		return date;
	}

	public void setDate(Date date) {
		this.date = date;
	}

	public String getComments() {
		return comments;
	}

	public void setComments(String comments) {
		this.comments = comments;
	}

	public String getProject() {
		return project;
	}

	public void setProject(String project) {
		this.project = project;
	}

	public String getTask() {
		return task;
	}

	public void setTask(String task) {
		this.task = task;
	}

	public Double getWorkHours() {
		return workHours;
	}

	public void setWorkHours(Double workHours) {
		this.workHours = workHours;
	}

	@Override
	public int compareTo(WorkHourRow whr) {
		return this.getDate().compareTo(whr.getDate());
	}

	@Override
	public int hashCode() {
		return Objects.hash(comments, date, project, task, workHours);
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		WorkHourRow other = (WorkHourRow) obj;
		return Objects.equals(comments, other.comments) && Objects.equals(date, other.date)
				&& Objects.equals(project, other.project) && Objects.equals(task, other.task)
				&& Objects.equals(workHours, other.workHours);
	}

	@Override
	public String toString() {
		return "WorkHourRow [date=" + date + ", comments=" + comments + ", project=" + project + ", task=" + task
				+ ", workHours=" + workHours + "]";
	}
}
