import aspose.words as aw
import datetime
import io
import os
import aspose 
from pathlib import Path

def do_comparison(first_doc, second_doc, color, reject_highlights):
    options = aw.comparing.CompareOptions()
    options.granularity = aw.comparing.Granularity.CHAR_LEVEL

    if (first_doc.revisions.count == 0 and second_doc.revisions.count == 0):
        first_doc.compare(second_doc, "authorName", datetime.datetime.today(), options)
    else:
        print("revisions in one of the docs, can't compare")
        exit()

    to_reject = []

    for r in first_doc.revisions:
        try:
            if reject_highlights and r.group.text == "Highlight":
                to_reject.append(r)
                continue
            run = r.parent_node.as_run()
            run.font.highlight_color = color
        except Exception as e:
            # print(f"Revision type: {r.revision_type}, on a node of type \"{r.parent_node.node_type}\"")
            # print(f"\tChanged text: \"{r.parent_node.get_text()}\"")
            print(e)

    for change in to_reject:
        change.reject()


def main(): 
    lic = aw.License()

    # Try to set license from the stream.
    try:
        with io.FileIO("aspose.lic") as stream:
            lic.set_license(stream)
        print("License set successfully.")
    except RuntimeError as err:
        print("\nThere was an error setting the license:", err)

    Path("markups").mkdir(parents=True, exist_ok=True)

    # base_exam = aw.Document("base_exam.docx")
    answer_key = aw.Document("exam_answer_key.docx")

    directory = "submissions"
    for name in os.listdir(directory):
        submission = aw.Document(os.path.join(directory, name))

        # submission_clone = submission.clone().as_document()
        # answer_key_clone = answer_key.clone().as_document()
        # do_comparison(submission_clone, answer_key_clone, aspose.pydrawing.Color.yellow, False)

        # submission_clone.save(os.path.join("markups", name + "_markup_v1.docx"))

        submission_clone2 = submission.clone().as_document()
        submission_clone2.accept_all_revisions()
        answer_key_clone2 = answer_key.clone().as_document()
        do_comparison(answer_key_clone2, submission_clone2, aspose.pydrawing.Color.cyan, False)
        answer_key_clone2.save(os.path.join("markups", name + "_markup.docx"))


        # submission_clone3 = submission.clone().as_document()
        # base_exam_clone3 = base_exam.clone().as_document()
        # answer_key_clone3 = answer_key.clone().as_document()

        # do_comparison(base_exam_clone3, submission_clone3, aspose.pydrawing.Color.yellow, False)

        # base_exam_clone3.revisions.accept_all()
        # base_exam_clone3.save("diff.docx")

        # first_diff = aw.Document("diff.docx")

        # do_comparison(first_diff, answer_key_clone3, aspose.pydrawing.Color.cyan, True)

        # first_diff.save(os.path.join("markups", name + "_markup_v3.docx"))
        # os.remove("diff.docx")


if __name__=="__main__": 
    main() 
